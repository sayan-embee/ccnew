// <copyright file="RecallSentNotificationsController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Graph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RecallSentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.DataRecallQueue;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.PrepareToRecallQueue;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Controller for the recalling sent notification data.
    /// </summary>
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    [Route("api/recallSentNotifications")]
    public class RecallSentNotificationsController : ControllerBase
    {
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly IRecallSentNotificationDataRepository recallSentNotificationDataRepository;
        private readonly IPrepareToRecallQueue prepareToRecallQueue;
        private readonly IDataRecallQueue dataRecallQueue;
        private readonly double forceCompleteMessageDelayInSeconds;
        private readonly ILogger<RecallSentNotificationsController> logger;
        private readonly IGraphServiceFactoryByTenant graphServiceFactoryByTenant;
        private readonly ITeamDataRepository teamDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="RecallSentNotificationsController"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">Notification data repository service that deals with the table storage in azure.</param>
        /// <param name="recallSentNotificationDataRepository">Sent notification data repository.</param>
        /// <param name="prepareToRecallQueue">The service bus queue for preparing to send notifications.</param>
        /// <param name="dataRecallQueue">The service bus queue for the data queue.</param>
        /// <param name="dataQueueMessageOptions">The options for the data queue messages.</param>
        /// <param name="loggerFactory">The logger factory.</param>
        /// <param name="graphServiceFactoryByTenant">graphServiceFactoryByTenant.</param>
        /// <param name="teamDataRepository">Team data repository instance.</param>
        public RecallSentNotificationsController(
            INotificationDataRepository notificationDataRepository,
            IRecallSentNotificationDataRepository recallSentNotificationDataRepository,
            IPrepareToRecallQueue prepareToRecallQueue,
            IDataRecallQueue dataRecallQueue,
            IOptions<DataQueueMessageOptions> dataQueueMessageOptions,
            ILoggerFactory loggerFactory,
            IGraphServiceFactoryByTenant graphServiceFactoryByTenant,
            ITeamDataRepository teamDataRepository)
        {
            if (dataQueueMessageOptions is null)
            {
                throw new ArgumentNullException(nameof(dataQueueMessageOptions));
            }

            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.recallSentNotificationDataRepository = recallSentNotificationDataRepository ?? throw new ArgumentNullException(nameof(recallSentNotificationDataRepository));
            this.prepareToRecallQueue = prepareToRecallQueue ?? throw new ArgumentNullException(nameof(prepareToRecallQueue));
            this.dataRecallQueue = dataRecallQueue ?? throw new ArgumentNullException(nameof(dataRecallQueue));
            this.forceCompleteMessageDelayInSeconds = dataQueueMessageOptions.Value.ForceCompleteMessageDelayInSeconds;
            this.logger = loggerFactory?.CreateLogger<RecallSentNotificationsController>() ?? throw new ArgumentNullException(nameof(loggerFactory));
            this.graphServiceFactoryByTenant = graphServiceFactoryByTenant ?? throw new ArgumentNullException(nameof(graphServiceFactoryByTenant));
            this.teamDataRepository = teamDataRepository ?? throw new ArgumentNullException(nameof(teamDataRepository));
        }

        /// <summary>
        /// Send a notification, which turns a draft to be a sent notification.
        /// </summary>
        /// <param name="sentNotification">An instance of <see cref="SentNotification"/> class.</param>
        /// <returns>The result of an action method.</returns>
        [HttpPost]
        public async Task<IActionResult> CreateRecallNotificationAsync(
            [FromBody] SentNotification sentNotification)
        {
            if (sentNotification == null)
            {
                throw new ArgumentNullException(nameof(SentNotification));
            }

            var sendNotificationDataEntity = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.SentNotificationsPartition,
                sentNotification.Id);
            if (sendNotificationDataEntity == null)
            {
                return this.NotFound($"Recall Sent notification, Id: {sentNotification.Id}, could not be found.");
            }

            var newSentNotificationId =
                await this.notificationDataRepository.MoveSentToRecallPartitionAsync(sendNotificationDataEntity);

            // Ensure the data table needed by the Azure Functions to send the notifications exist in Azure storage.
            await this.recallSentNotificationDataRepository.EnsureRecallSentNotificationDataTableExistsAsync();

            var prepareToRecallQueueMessageContent = new PrepareToRecallQueueMessageContent
            {
                NotificationId = newSentNotificationId,
            };
            await this.prepareToRecallQueue.SendAsync(prepareToRecallQueueMessageContent);

            // Send a "force complete" message to the data queue with a delay to ensure that
            // the notification will be marked as complete no matter the counts
            var forceCompleteDataRecallQueueMessageContent = new DataRecallQueueMessageContent
            {
                NotificationId = newSentNotificationId,
                ForceMessageComplete = true,
            };
            await this.dataRecallQueue.SendDelayedAsync(
                forceCompleteDataRecallQueueMessageContent,
                this.forceCompleteMessageDelayInSeconds);

            return this.Ok();
        }

        /// <summary>
        /// Get recall notifications.
        /// </summary>
        /// <returns>A list of <see cref="RecallSentNotificationSummary"/> instances.</returns>
        [HttpGet]
        public async Task<ActionResult<IEnumerable<RecallSentNotificationSummary>>> GetAllRecallSentNotificationsAsync()
        {
            var notificationEntities = await this.notificationDataRepository.GetAllRecallSentNotificationsAsync();

            var result = new List<RecallSentNotificationSummary>();
            this.logger.LogTrace($"GetAllRecallSentNotificationsAsync");
            foreach (var notificationEntity in notificationEntities)
            {
                this.logger.LogTrace($"GetAllRecallSentNotificationsAsync Notifiation Id: {notificationEntity.Id}");
                var summary = new RecallSentNotificationSummary
                {
                    Id = notificationEntity.Id,
                    Title = notificationEntity.Title,
                    CreatedDateTime = notificationEntity.CreatedDate,
                    SentDate = notificationEntity.RecalledDate,
                    Succeeded = notificationEntity.Succeeded,
                    Failed = notificationEntity.Failed,
                    Unknown = this.GetUnknownCount(notificationEntity),
                    TotalMessageCount = notificationEntity.TotalMessageCount,
                    SendingStartedDate = notificationEntity.SendingStartedDate,
                    Status = notificationEntity.GetRecalledStatus(),
                    TenantName = notificationEntity.TenantName,
                };
                result.Add(summary);
            }

            return result;
        }

        /// <summary>
        /// Get a recall sent notification by Id.
        /// </summary>
        /// <param name="id">Id of the requested recall sent notification.</param>
        /// <returns>Required sent notification.</returns>
        [HttpGet("{id}")]
        public async Task<IActionResult> GetRecallNotificationByIdAsync(string id)
        {
            if (id == null)
            {
                throw new ArgumentNullException(nameof(id));
            }

            var notificationEntity = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.RecallSentNotificationsPartition,
                id);
            if (notificationEntity == null)
            {
                return this.NotFound();
            }

            var grpMemberService = this.graphServiceFactoryByTenant.GetGroupsService(notificationEntity.TenantId);
            var groupNames = await grpMemberService.
                GetByIdsAsync(notificationEntity.Groups).
                Select(x => x.DisplayName).
                ToListAsync();

            /* var groupNames = await this.groupsService.
                GetByIdsAsync(notificationEntity.Groups).
                Select(x => x.DisplayName).
                ToListAsync();
            */

            // var userId = this.HttpContext.User.FindFirstValue(Common.Constants.ClaimTypeUserId);
            // var userNotificationDownload = await this.exportDataRepository.GetAsync(userId, id);
            var result = new
            {
                Id = notificationEntity.Id,
                Title = notificationEntity.Title,
                ImageLink = notificationEntity.ImageLink,
                Summary = notificationEntity.Summary,
                Author = notificationEntity.Author,
                ButtonTitle = notificationEntity.ButtonTitle,
                ButtonLink = notificationEntity.ButtonLink,
                Buttons = notificationEntity.Buttons,
                IsScheduled = notificationEntity.IsScheduled,
                IsImportant = notificationEntity.IsImportant,
                CreatedDateTime = notificationEntity.CreatedDate,
                SentDate = notificationEntity.SentDate,
                Succeeded = notificationEntity.Succeeded,
                Failed = notificationEntity.Failed,
                Unknown = this.GetUnknownCount(notificationEntity),
                TeamNames = await this.teamDataRepository.GetTeamNamesByIdsAsync(notificationEntity.Teams),
                RosterNames = await this.teamDataRepository.GetTeamNamesByIdsAsync(notificationEntity.Rosters),
                GroupNames = groupNames,
                AllUsers = notificationEntity.AllUsers,
                SendingStartedDate = notificationEntity.SendingStartedDate,
                ErrorMessage = notificationEntity.ErrorMessage,
                WarningMessage = notificationEntity.WarningMessage,
                CanDownload = true,
                SendingCompleted = notificationEntity.IsCompleted(),
                TemplateType = notificationEntity.TemplateType,
                TenantId = notificationEntity.TenantId,
                SendTypeId = notificationEntity.SendTypeId,
                AdaptiveCardContent = notificationEntity.AdaptiveCardContent,
                EmailBody = notificationEntity.EmailBody,
                EmailTitle = notificationEntity.Title,
                CsvLink = notificationEntity.CsvLink,
                AdditionalFileLink = notificationEntity.AdditionalFileLink,
                TenantName = notificationEntity.TenantName,
                AuthorTeamId = notificationEntity.AuthorTeamId,
                AuthorTeamName = notificationEntity.AuthorTeamName,
                AuthorChannelId = notificationEntity.AuthorChannelId,
                AuthorChannelName = notificationEntity.AuthorChannelName,
                RecalledDate = notificationEntity.RecalledDate,
                TotalMessageCount = notificationEntity.TotalMessageCount,
                TotalReadReceipt = 0,
            };

            return this.Ok(result);
        }


        /// <summary>
        /// Get a recall sent notification by Id.
        /// </summary>
        /// <param name="id">Id of the requested recall sent notification.</param>
        /// <returns>Required sent notification.</returns>
        [HttpGet("download/{id}")]
        public async Task<IActionResult> GetDownloadNotificationByIdAsync(string id)
        {
            if (id == null)
            {
                throw new ArgumentNullException(nameof(id));
            }

            var notificationEntity = await this.recallSentNotificationDataRepository.GetAllAsync(partition: id);
            if (notificationEntity == null)
            {
                return this.NotFound();
            }
            var recallNotificationDataEntity = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.RecallSentNotificationsPartition,
                id);
            if (recallNotificationDataEntity != null)
            {
                foreach (var entity in notificationEntity)
                {
                    entity.TenantName = recallNotificationDataEntity.TenantName;
                }
            }
            return this.Ok(notificationEntity);
        }

        private int? GetUnknownCount(NotificationDataEntity notificationEntity)
        {
            var unknown = notificationEntity.Unknown;

            // In CC v2, the number of throttled recipients are counted and saved in NotificationDataEntity.Unknown property.
            // However, CC v1 saved the number of throttled recipients in NotificationDataEntity.Throttled property.
            // In order to make it backward compatible, we add the throttled number to the unknown variable.
            var throttled = notificationEntity.Throttled;
            if (throttled > 0)
            {
                unknown += throttled;
            }

            return unknown > 0 ? unknown : (int?)null;
        }
    }
}

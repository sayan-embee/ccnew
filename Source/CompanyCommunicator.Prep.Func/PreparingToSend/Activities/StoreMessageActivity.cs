// <copyright file="StoreMessageActivity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System;
    using System.Text.Json;
    using System.Text.Json.Serialization;
    using System.Threading.Tasks;
    using Microsoft.Azure.Functions.Worker;
    using Microsoft.DurableTask;
    using Microsoft.DurableTask.Client;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard;

    /// <summary>
    /// Stores the message in sending notification data table.
    /// </summary>
    public class StoreMessageActivity
    {
        private readonly ISendingNotificationDataRepository sendingNotificationDataRepository;
        private readonly AdaptiveCardCreator adaptiveCardCreator;

        /// <summary>
        /// Initializes a new instance of the <see cref="StoreMessageActivity"/> class.
        /// </summary>
        /// <param name="notificationRepo">Sending notification data repository.</param>
        /// <param name="cardCreator">The adaptive card creator.</param>
        public StoreMessageActivity(
            ISendingNotificationDataRepository notificationRepo,
            AdaptiveCardCreator cardCreator)
        {
            this.sendingNotificationDataRepository = notificationRepo ?? throw new ArgumentNullException(nameof(notificationRepo));
            this.adaptiveCardCreator = cardCreator ?? throw new ArgumentNullException(nameof(cardCreator));
        }

        /// <summary>
        /// Stores the message in sending notification data table.
        /// </summary>
        /// <param name="notification">A notification to be sent to recipients.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        [Function(FunctionNames.StoreMessageActivity)]
        public async Task RunAsync(
            [ActivityTrigger] NotificationDataEntity notification)
        {
            if (notification == null)
            {
                throw new ArgumentNullException(nameof(notification));
            }

            var serializedContent = this.adaptiveCardCreator.CreateAdaptiveCard(notification).ToJson();
            var serializeobject = JsonSerializer.Serialize(notification.AdaptiveCardContent);

            var sendingNotification = new SendingNotificationDataEntity
            {
                PartitionKey = NotificationDataTableNames.SendingNotificationsPartition,
                RowKey = notification.RowKey,
                NotificationId = notification.Id,
                Content = notification.AdaptiveCardContent, // serializedContent,
                EmailBody = notification.EmailBody,
                CreatedBy = notification.CreatedBy,
                TenantId = notification.TenantId,
                EmailTitle = notification.EmailTitle,
                SendTypeId = notification.SendTypeId,
                TemplateType = notification.TemplateType,
                Title = notification.Title,
                ImageLink = notification.ImageLink,
                CsvLink = notification.CsvLink,
                AdditionalFileLink = notification.AdditionalFileLink,
                TenantName = notification.TenantName,
                AuthorTeamId = notification.AuthorTeamId,
                AuthorTeamName = notification.AuthorTeamName,
                AuthorChannelId = notification.AuthorChannelId,
                AuthorChannelName = notification.AuthorChannelName,
            };

            await this.sendingNotificationDataRepository.CreateOrUpdateAsync(sendingNotification);
        }
    }
}

// <copyright file="ReadReceiptNotificationController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;

    /// <summary>
    /// Controller for the read receipt of notification data.
    /// </summary>
    [Route("api/ReadReceiptNotification")]
    [ApiController]
    public class ReadReceiptNotificationController : ControllerBase
    {
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly ILogger<ReadReceiptNotificationController> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="ReadReceiptNotificationController"/> class.
        /// </summary>
        /// <param name="sentNotificationDataRepository">Sent notification data repository.</param>
        /// <param name="notificationDataRepository">Notification data repository service that deals with the table storage in azure.</param>
        /// <param name="loggerFactory">The logger factory.</param>
        public ReadReceiptNotificationController(
            ISentNotificationDataRepository sentNotificationDataRepository,
            INotificationDataRepository notificationDataRepository,
            ILoggerFactory loggerFactory)
        {
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentNullException(nameof(sentNotificationDataRepository));
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.logger = loggerFactory?.CreateLogger<ReadReceiptNotificationController>() ?? throw new ArgumentNullException(nameof(loggerFactory));
        }

        /// <summary>
        /// Get a sent notification by Id.
        /// </summary>
        /// <param name="id">Id of the requested sent notification.</param>
        /// <param name="rowKey">RowKey to the table of the requested sent notification.</param>
        /// <returns>Required sent notification.</returns>
        [HttpGet]
        [Route("view")]
        public async Task<IActionResult> GetSentNotificationByIdAsync(string id, string rowKey)
        {
            try
            {
                if (id == null || rowKey == null)
                {
                    // throw new ArgumentNullException(nameof(id));
                    return this.Ok("Id not found");
                }
                else
                {
                    var sentNotificationEntity = await this.sentNotificationDataRepository.GetAsync(id, rowKey);
                    if (sentNotificationEntity == null)
                    {
                        this.logger.LogInformation($"Unable to get Sent Notification Data for Partition Key = {id} and Row Key = {rowKey}.");
                        return this.NotFound();
                    }

                    // this.logger.LogInformation($"Unable to get Sent Notification Data for Partition Key ={id} and Row Key={rowKey} and read receipt = {sentNotificationEntity.ReadReceipt}.");
                    if (sentNotificationEntity.ReadReceipt == 0)
                    {
                        sentNotificationEntity.ReadReceipt = 1;
                        sentNotificationEntity.ReadReceiptDate = DateTime.UtcNow;
                        await this.sentNotificationDataRepository.InsertOrMergeAsync(sentNotificationEntity);
                        this.logger.LogInformation($"Read receipt done for Sent Notification Data for Partition Key = {id} and Row Key = {rowKey}.");

                        var partitionKey = "SentNotifications";
                        var notificationEntity = await this.notificationDataRepository.GetAsync(partitionKey, id);
                        if (notificationEntity == null)
                        {
                            this.logger.LogInformation($"Unable to get Sent Notification Data for Partition Key = {id} and Row Key = {rowKey}.");
                            return this.NotFound();
                        }

                        notificationEntity.TotalReadReceipt = notificationEntity.TotalReadReceipt + 1;
                        await this.notificationDataRepository.InsertOrMergeAsync(notificationEntity);
                        this.logger.LogInformation($"Total Read receipt done for Sent Notification Data for Partition Key = {id} and Row Key = {rowKey}.");
                    }

                    return this.Ok();
                }
            }
            catch (Exception exception)
            {
                this.logger.LogError(exception, $"Failed to get log read receipt. Error message: {exception.Message}.");
                return this.Problem(exception.Message);
            }
        }
    }
}
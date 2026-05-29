// <copyright file="UpdateRecallNotificationDataService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Data.Func.Services.NotificationDataServices
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;

    /// <summary>
    /// Service to update notification data.
    /// </summary>
    public class UpdateRecallNotificationDataService
    {
        private readonly INotificationDataRepository notificationDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="UpdateRecallNotificationDataService"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">The notification data repository.</param>
        public UpdateRecallNotificationDataService(
            INotificationDataRepository notificationDataRepository)
        {
            this.notificationDataRepository = notificationDataRepository;
        }

        /// <summary>
        /// Updates the notification totals with the given information and results.
        /// </summary>
        /// <param name="notificationId">The notification ID.</param>
        /// <param name="shouldForceCompleteNotification">Flag to indicate if the notification should
        /// be forced to be marked as completed.</param>
        /// <param name="totalExpectedNotificationCount">The total expected count of notifications to be sent.</param>
        /// <param name="aggregatedRecallSentNotificationDataResults">The current aggregated results for
        /// the sent notifications.</param>
        /// <param name="log">The logger.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public async Task<UpdateRecallNotificationDataEntity> UpdateRecallNotificationDataAsync(
            string notificationId,
            bool shouldForceCompleteNotification,
            int totalExpectedNotificationCount,
            AggregatedRecallSentNotificationDataResults aggregatedRecallSentNotificationDataResults,
            ILogger log)
        {
            try
            {
                var currentTotalNotificationCount = aggregatedRecallSentNotificationDataResults.CurrentTotalNotificationCount;
                var succeededCount = aggregatedRecallSentNotificationDataResults.SucceededCount;
                var failedCount = aggregatedRecallSentNotificationDataResults.FailedCount;
                var throttledCount = aggregatedRecallSentNotificationDataResults.ThrottledCount;
                var recipientNotFoundCount = aggregatedRecallSentNotificationDataResults.RecipientNotFoundCount;
                var lastSentDate = aggregatedRecallSentNotificationDataResults.LastSentDate;

                // Create the general update.
                var notificationDataEntityUpdate = new UpdateRecallNotificationDataEntity
                {
                    PartitionKey = NotificationDataTableNames.RecallSentNotificationsPartition,
                    RowKey = notificationId,
                    Succeeded = succeededCount,
                    Failed = failedCount,
                    RecipientNotFound = recipientNotFoundCount,
                    Throttled = throttledCount,
                };

                // If it should be marked as complete, set the other values accordingly.
                if (currentTotalNotificationCount >= totalExpectedNotificationCount
                    || shouldForceCompleteNotification)
                {
                    // Update the status to Sent.
                    notificationDataEntityUpdate.Status = NotificationStatus.Recalled.ToString();

                    if (currentTotalNotificationCount >= totalExpectedNotificationCount)
                    {
                        // If the message is being completed because all messages have been accounted for,
                        // then make sure the unknown count is 0 and update the sent date with the date
                        // of the last sent message.
                        notificationDataEntityUpdate.Unknown = 0;
                        notificationDataEntityUpdate.RecalledDate = lastSentDate ?? DateTime.UtcNow;
                    }
                    else if (shouldForceCompleteNotification)
                    {
                        // If the message is being completed, not because all messages have been accounted for,
                        // but because the trigger is coming from the delayed Service Bus message that ensures that the
                        // notification will eventually be marked as complete, then update the unknown count of messages
                        // not accounted for and update the sent date to the current time.
                        var countDifference = totalExpectedNotificationCount - currentTotalNotificationCount;

                        // This count must stay 0 or above.
                        var unknownCount = countDifference >= 0 ? countDifference : 0;

                        notificationDataEntityUpdate.Unknown = unknownCount;
                        notificationDataEntityUpdate.RecalledDate = DateTime.UtcNow;
                    }

                    if (notificationDataEntityUpdate.Status == NotificationStatus.Recalled.ToString())
                    {
                        // All the messages has been sent now update the sent notification IsRecalled to true
                        var sendNotificationDataEntity = await this.notificationDataRepository.GetAsync(
                        NotificationDataTableNames.SentNotificationsPartition,
                        notificationId);

                        if (sendNotificationDataEntity != null)
                        {
                            sendNotificationDataEntity.IsRecalled = true;
                            await this.notificationDataRepository.InsertOrMergeAsync(sendNotificationDataEntity);
                        }
                    }
                }

                var operation = TableOperation.InsertOrMerge(notificationDataEntityUpdate);
                await this.notificationDataRepository.Table.ExecuteAsync(operation);

                return notificationDataEntityUpdate;
            }
            catch (Exception e)
            {
                var errorMessage = $"{e.GetType()}: {e.Message}";
                log.LogError(e, $"ERROR: {errorMessage}");
                throw;
            }
        }
    }
}

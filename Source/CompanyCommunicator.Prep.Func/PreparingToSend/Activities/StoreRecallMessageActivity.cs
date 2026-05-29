// <copyright file="StoreRecallMessageActivity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.Json;
    using System.Text.Json.Serialization;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Azure.Functions.Worker;
    using Microsoft.DurableTask;
    using Microsoft.DurableTask.Client;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RecallSentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend.Extensions;

    /// <summary>
    /// Stores the message in sending notification data table.
    /// </summary>
    public class StoreRecallMessageActivity
    {
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly IRecallSentNotificationDataRepository recallSentNotificationDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="StoreRecallMessageActivity"/> class.
        /// </summary>
        /// <param name="sentNotificationDataRepository">Sending notification data repository.</param>
        /// <param name="recallSentNotificationDataRepository">recallSentNotificationDataRepository.</param>
        public StoreRecallMessageActivity(
            ISentNotificationDataRepository sentNotificationDataRepository,
            IRecallSentNotificationDataRepository recallSentNotificationDataRepository)
        {
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentNullException(nameof(sentNotificationDataRepository));
            this.recallSentNotificationDataRepository = recallSentNotificationDataRepository ?? throw new ArgumentNullException(nameof(recallSentNotificationDataRepository));
        }

        /// <summary>
        /// Stores the message in sending notification data table.
        /// </summary>
        /// <param name="notification">A notification to be sent to recipients.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        [Function(FunctionNames.StoreRecallMessageActivity)]
        public async Task RunAsync(
            [ActivityTrigger] NotificationDataEntity notification)
        {
            if (notification == null)
            {
                throw new ArgumentNullException(nameof(notification));
            }

            // Get sent notification records
            string filter = TableQuery.GenerateFilterCondition("DeliveryStatus", QueryComparisons.Equal, "Succeeded");
            var recipients = await this.sentNotificationDataRepository.GetWithFilterAsync(filter, notification.Id);
            if (recipients == null)
            {
                throw new ArgumentNullException(nameof(recipients));
            }

            if (recipients != null)
            {
                var recallNotifications = recipients.Select(
                 sentNotification => sentNotification.CreateInitialRecallSentNotificationDataEntity(partitionKey: notification.Id));

                await this.recallSentNotificationDataRepository.BatchInsertOrMergeAsync(recallNotifications);
            }
        }
    }
}

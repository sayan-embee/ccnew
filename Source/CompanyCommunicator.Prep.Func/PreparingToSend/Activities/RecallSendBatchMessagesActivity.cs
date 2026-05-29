// <copyright file="RecallSendBatchMessagesActivity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Functions.Worker;
    using Microsoft.DurableTask;
    using Microsoft.DurableTask.Client;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RecallSentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.SendRecallQueue;

    /// <summary>
    /// Recall Sends batch messages to Send Queue.
    /// </summary>
    public class RecallSendBatchMessagesActivity
    {
        private readonly ISendRecallQueue sendRecallQueue;

        /// <summary>
        /// Initializes a new instance of the <see cref="RecallSendBatchMessagesActivity"/> class.
        /// </summary>
        /// <param name="sendRecallQueue">Send queue service.</param>
        public RecallSendBatchMessagesActivity(
            ISendRecallQueue sendRecallQueue)
        {
            this.sendRecallQueue = sendRecallQueue ?? throw new ArgumentNullException(nameof(sendRecallQueue));
        }

        /// <summary>
        /// Sends batch messages to Send Queue.
        /// </summary>
        /// <param name="input">Input.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [Function(FunctionNames.RecallSendBatchMessagesActivity)]
        public async Task RunAsync(
            [ActivityTrigger](NotificationDataEntity notification, List<RecallSentNotificationDataEntity> batch) input)
        {
            if (input.notification == null)
            {
                throw new ArgumentNullException(nameof(input.notification));
            }

            if (input.batch == null)
            {
                throw new ArgumentNullException(nameof(input.batch));
            }

            var messageBatch = input.batch.Select(
                recipient =>
                {
                    return new SendRecallQueueMessageContent()
                    {
                        NotificationId = input.notification.Id,
                        RecipientData = recipient,
                        MessageTitle = input.notification.Title,

                        // RecipientData = this.ConvertToRecipientData(recipient, input.notification.TenantId),
                    };
                });

            await this.sendRecallQueue.SendAsync(messageBatch);
        }
    }
}

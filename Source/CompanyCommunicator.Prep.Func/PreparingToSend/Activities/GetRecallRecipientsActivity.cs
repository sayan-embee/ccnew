// <copyright file="GetRecallRecipientsActivity.cs" company="Microsoft">
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

    /// <summary>
    /// Reads all the recipients from recall notification table.
    /// </summary>
    public class GetRecallRecipientsActivity
    {
        private readonly IRecallSentNotificationDataRepository recallSentNotificationDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetRecallRecipientsActivity"/> class.
        /// </summary>
        /// <param name="recallSentNotificationDataRepository">The recall sent notification data repository.</param>
        public GetRecallRecipientsActivity(IRecallSentNotificationDataRepository recallSentNotificationDataRepository)
        {
            this.recallSentNotificationDataRepository = recallSentNotificationDataRepository ?? throw new ArgumentNullException(nameof(recallSentNotificationDataRepository));
        }

        /// <summary>
        /// Reads all the recipients from Sent notification table.
        /// </summary>
        /// <param name="notification">notification.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [Function(FunctionNames.GetRecipientsRecallActivity)]
        public async Task<IEnumerable<RecallSentNotificationDataEntity>> GetRecipientsRecallAsync([ActivityTrigger] NotificationDataEntity notification)
        {
            if (notification == null)
            {
                throw new ArgumentNullException(nameof(notification));
            }

            var recipients = await this.recallSentNotificationDataRepository.GetAllAsync(notification.Id);
            return recipients;
        }

        /// <summary>
        /// Reads all the recipients from Sent notification table who do not have conversation details.
        /// </summary>
        /// <param name="notification">notification.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [Function(FunctionNames.GetPendingRecipientsRecallActivity)]
        public async Task<IEnumerable<RecallSentNotificationDataEntity>> GetPendingRecipientsRecallAsync([ActivityTrigger] NotificationDataEntity notification)
        {
            if (notification == null)
            {
                throw new ArgumentNullException(nameof(notification));
            }

            var recipients = await this.recallSentNotificationDataRepository.GetAllAsync(notification.Id);
            return recipients.Where(recipient => string.IsNullOrEmpty(recipient.ConversationId));
        }
    }
}

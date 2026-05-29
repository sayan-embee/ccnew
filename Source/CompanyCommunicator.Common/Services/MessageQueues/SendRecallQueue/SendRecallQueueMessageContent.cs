// <copyright file="SendRecallQueueMessageContent.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.SendRecallQueue
{
    using System.Collections.Generic;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RecallSentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;

    /// <summary>
    /// Azure service bus send queue message content class.
    /// </summary>
    public class SendRecallQueueMessageContent
    {
        /// <summary>
        /// Gets or sets the notification id value.
        /// </summary>
        public string NotificationId { get; set; }

        /// <summary>
        /// Gets or sets the information about the recipient. This
        /// holds enough information for the Azure Function to send this
        /// recipient a notification.
        /// </summary>
        public RecallSentNotificationDataEntity RecipientData { get; set; }

        /// <summary>
        /// Gets or sets the Message Title value.
        /// </summary>
        public string MessageTitle { get; set; }
    }
}

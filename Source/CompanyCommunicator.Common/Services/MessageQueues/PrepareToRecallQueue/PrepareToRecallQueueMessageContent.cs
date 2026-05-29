// <copyright file="PrepareToRecallQueueMessageContent.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.PrepareToRecallQueue
{
    /// <summary>
    /// Azure service bus prepare to send queue message content class.
    /// </summary>
    public class PrepareToRecallQueueMessageContent
    {
        /// <summary>
        /// Gets or sets the notification id value.
        /// </summary>
        public string NotificationId { get; set; }
    }
}

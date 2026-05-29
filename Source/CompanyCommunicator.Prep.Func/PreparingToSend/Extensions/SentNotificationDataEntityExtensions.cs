// <copyright file="SentNotificationDataEntityExtensions.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend.Extensions
{
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RecallSentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;

    /// <summary>
    /// Extension methods for the SentNotificationDataEntity class.
    /// </summary>
    public static class SentNotificationDataEntityExtensions
    {
        /// <summary>
        /// Creates a SentNotificationDataEntity in an initialized state from the given UserDataEntity
        /// and partition key.
        /// Makes sure to set the correct recipient type for having been created from a UserDataEntity.
        /// </summary>
        /// <param name="sentNotificationDataEntity">The user data entity.</param>
        /// <param name="partitionKey">The partition key.</param>
        /// <returns>The sent notification data entity.</returns>
        public static RecallSentNotificationDataEntity CreateInitialRecallSentNotificationDataEntity(
            this SentNotificationDataEntity sentNotificationDataEntity,
            string partitionKey)
        {
            return new RecallSentNotificationDataEntity
            {
                PartitionKey = partitionKey,
                RowKey = sentNotificationDataEntity.RowKey,
                ActivityId = sentNotificationDataEntity.ActivityId,
                ConversationId = sentNotificationDataEntity.ConversationId,
                TenantId = sentNotificationDataEntity.TenantId,
                UserId = sentNotificationDataEntity.UserId,
                ServiceUrl = sentNotificationDataEntity.ServiceUrl,
                Email = sentNotificationDataEntity.Email,
                Name = sentNotificationDataEntity.Name,
                Upn = sentNotificationDataEntity.Upn,
                RecipientId = sentNotificationDataEntity.RecipientId,
                RecipientType = sentNotificationDataEntity.RecipientType,
            };
        }
    }
}

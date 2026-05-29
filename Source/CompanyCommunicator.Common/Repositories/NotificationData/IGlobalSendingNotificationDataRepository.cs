// <copyright file="IGlobalSendingNotificationDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData
{
    using System.Threading.Tasks;

    /// <summary>
    /// interface for Global Sending Notification Data Repository.
    /// </summary>
    public interface IGlobalSendingNotificationDataRepository : IRepository<GlobalSendingNotificationDataEntity>
    {
        /// <summary>
        /// Gets the entity that holds meta-data for all sending operations.
        /// </summary>
        /// <returns>The Global Sending Notification Data Entity.</returns>
        public Task<GlobalSendingNotificationDataEntity> GetGlobalSendingNotificationDataEntityAsync();

        /// <summary>
        /// GetGlobalSendingRecallNotificationDataEntityAsync.
        /// </summary>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        Task<GlobalSendingNotificationDataEntity> GetGlobalSendingRecallNotificationDataEntityAsync();

        /// <summary>
        /// Insert or merges the entity that holds meta-data for all sending operations. Partition Key and Row Key do not need to be
        /// set on the incoming entity.
        /// </summary>
        /// <param name="globalSendingNotificationDataEntity">Entity that holds meta-data for all sending operations. Partition Key and
        /// Row Key do not need to be set.</param>
        /// <returns>The Task.</returns>
        public Task SetGlobalSendingNotificationDataEntityAsync(GlobalSendingNotificationDataEntity globalSendingNotificationDataEntity);

        /// <summary>
        /// SetGlobalSendingRecallNotificationDataEntityAsync.
        /// </summary>
        /// <param name="globalSendingRecallNotificationDataEntity">globalSendingRecallNotificationDataEntity.</param>
        /// <returns>A <see cref="Task"/> representing the result of the asynchronous operation.</returns>
        Task SetGlobalSendingRecallNotificationDataEntityAsync(GlobalSendingNotificationDataEntity globalSendingRecallNotificationDataEntity);
    }
}

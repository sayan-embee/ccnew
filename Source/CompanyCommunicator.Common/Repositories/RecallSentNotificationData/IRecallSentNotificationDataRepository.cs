// <copyright file="IRecallSentNotificationDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RecallSentNotificationData
{
    using System.Threading.Tasks;

    /// <summary>
    /// Interface for Sent Notification data Repository.
    /// </summary>
    public interface IRecallSentNotificationDataRepository : IRepository<RecallSentNotificationDataEntity>
    {
        /// <summary>
        /// This method ensures the RecallSentNotificationData table is created in the storage.
        /// This method should be called before kicking off an Azure function that uses the RecallSentNotificationData table.
        /// Otherwise the app will crash.
        /// By design, Azure functions (in this app) do not create a table if it's absent.
        /// </summary>
        /// <returns>A <see cref="Task"/> representing the result of the asynchronous operation.</returns>
        Task EnsureRecallSentNotificationDataTableExistsAsync();
    }
}
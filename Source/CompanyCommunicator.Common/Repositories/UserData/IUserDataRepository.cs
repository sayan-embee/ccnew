// <copyright file="IUserDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData
{
    using System.Collections.Generic;
    using System.Threading.Tasks;

    /// <summary>
    /// Interface for User Data Repository.
    /// </summary>
    public interface IUserDataRepository : IRepository<UserDataEntity>
    {
        /// <summary>
        /// Get delta link.
        /// </summary>
        /// <returns>Delta link.</returns>
        public Task<string> GetDeltaLinkAsync();

        /// <summary>
        /// Get delta link.
        /// </summary>
        /// <param name="tenantId">Tenant Id.</param>
        /// <returns>Delta link.</returns>
        Task<string> GetDeltaLinkAsync(string tenantId);

        /// <summary>
        /// Sets delta link.
        /// </summary>
        /// <param name="deltaLink">delta link.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        public Task SetDeltaLinkAsync(string deltaLink);

        /// <summary>
        /// Sets delta link.
        /// </summary>
        /// <param name="deltaLink">delta link.</param>
        /// <param name="tenantId">tenant Id.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        Task SetDeltaLinkAsync(string deltaLink, string tenantId);

        /// <summary>
        /// GetUserDataEntityByTenantIdAsync.
        /// </summary>
        /// <param name="tenantId">TenantId.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        Task<IEnumerable<UserDataEntity>> GetUserDataEntityByTenantIdAsync(string tenantId);
    }
}

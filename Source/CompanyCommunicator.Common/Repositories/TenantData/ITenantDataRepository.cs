// <copyright file="ITenantDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>
namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TenantData
{
    using System.Collections.Generic;
    using System.Threading.Tasks;

    /// <summary>
    /// Interface for Tenant Data Repository.
    /// </summary>
    public interface ITenantDataRepository : IRepository<TenantDataEntity>
    {
        /// <summary>
        /// EnsureTenantDataTableExistsAsync.
        /// </summary>
        /// <returns>A <see cref="Task"/> representing the result of the asynchronous operation.</returns>
        public Task EnsureTenantDataTableExistsAsync();
    }
}

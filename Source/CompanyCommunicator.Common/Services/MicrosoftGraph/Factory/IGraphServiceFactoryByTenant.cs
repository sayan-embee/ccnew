// <copyright file="IGraphServiceFactoryByTenant.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph
{
    using Microsoft.Graph;

    /// <summary>
    /// IGraphServiceFactoryByTenant.
    /// </summary>
    public interface IGraphServiceFactoryByTenant
    {
        /// <summary>
        /// GetAppManagerService.
        /// </summary>
        /// <param name="tenantId">Tenant Id.</param>
        /// <returns>IAppManagerService.</returns>
        IAppManagerServiceByTenant GetAppManagerService(string tenantId);

        /// <summary>
        /// GetAuthenticatedGraphClient.
        /// </summary>
        /// <param name="tenantId">Tenant Id.</param>
        /// <returns>GraphServiceClient.</returns>
        GraphServiceClient GetAuthenticatedGraphClient(string tenantId);

        /// <summary>
        /// GetChatsService.
        /// </summary>
        /// <param name="tenantId">Tenant Id.</param>
        /// <returns>IChatsService.</returns>
        IChatsServiceByTenant GetChatsService(string tenantId);

        /// <summary>
        /// GetGroupMembersService.
        /// </summary>
        /// <param name="tenantId">Tenant Id.</param>
        /// <returns>IGroupMembersService.</returns>
        IGroupMembersServiceByTenant GetGroupMembersService(string tenantId);

        /// <summary>
        /// GetGroupsService.
        /// </summary>
        /// <param name="tenantId">Tenant Id.</param>
        /// <returns>IGroupsService.</returns>
        IGroupsServiceByTenant GetGroupsService(string tenantId);

        /// <summary>
        /// GetUsersService.
        /// </summary>
        /// <param name="tenantId">Tenant Id.</param>
        /// <returns>IUsersService.</returns>
        IUsersServiceByTenant GetUsersService(string tenantId);

        /// <summary>
        /// Creates an instance of <see cref="IAppCatalogService"/> implementation.
        /// </summary>
        /// <param name="tenantId">Tenant Id.</param>
        /// <returns>Returns an implementation of <see cref="IAppCatalogService"/>.</returns>
        public IAppCatalogServiceByTenant GetAppCatalogService(string tenantId);
    }
}
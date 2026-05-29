// <copyright file="GraphServiceFactoryByTenant.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph
{
    using System;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Graph;
    using Microsoft.Identity.Client;

    /// <summary>
    /// Graph Service Factory.
    /// </summary>
    public class GraphServiceFactoryByTenant : IGraphServiceFactoryByTenant
    {
        private readonly IConfiguration configuration;
        private IGraphServiceClient serviceClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphServiceFactoryByTenant"/> class.
        /// </summary>
        /// <param name="configuration">IConfiguration.</param>
        public GraphServiceFactoryByTenant(IConfiguration configuration)
        {
            this.configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
        }

        /// <inheritdoc/>
        public GraphServiceClient GetAuthenticatedGraphClient(string tenantId)
        {
            var authenticationProvider = this.CreateAuthorizationProvider(tenantId);
            return new GraphServiceClient(authenticationProvider);
        }

        /// <inheritdoc/>
        public IUsersServiceByTenant GetUsersService(string tenantId)
        {
            this.serviceClient = this.GetAuthenticatedGraphClient(tenantId);

            return new UsersServiceByTenant(this.serviceClient);
        }

        /// <inheritdoc/>
        public IGroupsServiceByTenant GetGroupsService(string tenantId)
        {
            this.serviceClient = this.GetAuthenticatedGraphClient(tenantId);
            return new GroupsServiceByTenant(this.serviceClient);
        }

        /// <inheritdoc/>
        public IGroupMembersServiceByTenant GetGroupMembersService(string tenantId)
        {
            this.serviceClient = this.GetAuthenticatedGraphClient(tenantId);
            return new GroupMembersServiceByTenant(this.serviceClient);
        }

        /// <inheritdoc/>
        public IChatsServiceByTenant GetChatsService(string tenantId)
        {
            this.serviceClient = this.GetAuthenticatedGraphClient(tenantId);
            return new ChatsServiceByTenant(this.serviceClient, this.GetAppManagerService(tenantId));
        }

        /// <inheritdoc/>
        public IAppManagerServiceByTenant GetAppManagerService(string tenantId)
        {
            this.serviceClient = this.GetAuthenticatedGraphClient(tenantId);
            return new AppManagerServiceByTenant(this.serviceClient);
        }

        /// <inheritdoc/>
        public IAppCatalogServiceByTenant GetAppCatalogService(string tenantId)
        {
            this.serviceClient = this.GetAuthenticatedGraphClient(tenantId);
            return new AppCatalogServiceByTenant(this.serviceClient);
        }

        private IAuthenticationProvider CreateAuthorizationProvider(string tenantId)
        {
            var clientId = this.configuration.GetValue<string>("AuthorAppId");
            var clientSecret = this.configuration.GetValue<string>("AuthorAppPassword");

            var cca = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithClientSecret(clientSecret)
                .Build();
            return new MsalAuthenticationProvider(cca);
        }
    }
}

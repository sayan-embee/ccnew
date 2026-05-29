// <copyright file="AppCatalogServiceByTenant.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph
{
    using System;
    using System.Linq;
    using System.Threading.Tasks;

    using Microsoft.Graph;

    /// <summary>
    /// Read information about the apps published in the Teams app store and organization's app catalog.
    /// </summary>
    internal class AppCatalogServiceByTenant : IAppCatalogServiceByTenant
    {
        private readonly IGraphServiceClient graphServiceClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="AppCatalogServiceByTenant"/> class.
        /// </summary>
        /// <param name="graphServiceClient">Graph service client.</param>
        internal AppCatalogServiceByTenant(IGraphServiceClient graphServiceClient)
        {
            this.graphServiceClient = graphServiceClient ?? throw new ArgumentNullException(nameof(graphServiceClient));
        }

        /// <inheritdoc/>
        public async Task<string> GetTeamsAppIdAsync(string externalId)
        {
            if (externalId == null)
            {
                throw new ArgumentNullException(nameof(externalId));
            }

            var apps = await this.graphServiceClient
                .AppCatalogs
                .TeamsApps
                .Request()
                .Header(Common.Constants.PermissionTypeKey, GraphPermissionType.Application.ToString())
                .Filter($"distributionMethod eq 'organization' and externalId eq '{externalId}'")
                .GetAsync();

            return apps?.FirstOrDefault()?.Id;
        }
    }
}

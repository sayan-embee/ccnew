// <copyright file="AppSettingsService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;

    /// <summary>
    /// App settings service implementation.
    /// </summary>
    public class AppSettingsService : IAppSettingsService
    {
        private readonly IAppConfigRepository repository;

        private string serviceUrl;
        private string userAppId;

        /// <summary>
        /// Initializes a new instance of the <see cref="AppSettingsService"/> class.
        /// </summary>
        /// <param name="repository">App configuration repository.</param>
        public AppSettingsService(IAppConfigRepository repository)
        {
            this.repository = repository ?? throw new ArgumentNullException(nameof(repository));
        }

        /// <inheritdoc/>
        public async Task<string> GetServiceUrlAsync(string tenantId)
        {
            // check in-memory cache.
            // if (!string.IsNullOrWhiteSpace(this.serviceUrl))
            // {
            //    return this.serviceUrl;
            // }

            // var appConfig = await this.repository.GetAsync(
            //    AppConfigTableName.SettingsPartition,
            //    AppConfigTableName.ServiceUrlRowKey);
            // this.serviceUrl = appConfig?.Value;
            var appConfig = await this.repository.GetAsync(
               AppConfigTableName.SettingsServiceUrlPartition,
               tenantId);
            this.serviceUrl = appConfig?.Value;
            return this.serviceUrl;
        }

        /// <inheritdoc/>
        public async Task<string> GetUserAppIdAsync(string tenantId)
        {
            // check in-memory cache.
            // if (!string.IsNullOrWhiteSpace(this.userAppId))
            // {
            //    return this.userAppId;
            // }

            // var appConfig = await this.repository.GetAsync(
            //    AppConfigTableName.SettingsPartition,
            //    AppConfigTableName.UserAppIdRowKey);

            // this.userAppId = appConfig?.Value;

            //// return "30e987d7-8895-469d-a708-dd5945e559e3"; //this.userAppId;
            // return this.userAppId;
            var appConfig = await this.repository.GetAsync(AppConfigTableName.SettingsUserAppIdPartition, tenantId);
            this.userAppId = appConfig?.Value;

            return this.userAppId;
        }

        /// <inheritdoc/>
        public async Task SetServiceUrlAsync(string serviceUrl, string tenantId)
        {
            if (string.IsNullOrWhiteSpace(serviceUrl))
            {
                throw new ArgumentNullException(nameof(serviceUrl));
            }

            if (string.IsNullOrWhiteSpace(tenantId))
            {
                throw new ArgumentNullException(nameof(tenantId));
            }

            var appConfig = new AppConfigEntity()
            {
                PartitionKey = AppConfigTableName.SettingsServiceUrlPartition,
                RowKey = tenantId,
                Value = serviceUrl,
            };

            await this.repository.InsertOrMergeAsync(appConfig);

            // Update in-memory cache.
            // this.serviceUrl = serviceUrl;
        }

        /// <inheritdoc/>
        public async Task SetUserAppIdAsync(string userAppId, string tenantId)
        {
            if (string.IsNullOrWhiteSpace(userAppId))
            {
                throw new ArgumentNullException(nameof(userAppId));
            }

            if (string.IsNullOrWhiteSpace(tenantId))
            {
                throw new ArgumentNullException(nameof(tenantId));
            }

            var appConfig = new AppConfigEntity()
            {
                PartitionKey = AppConfigTableName.SettingsUserAppIdPartition,
                RowKey = tenantId,
                Value = userAppId,
            };

            await this.repository.InsertOrMergeAsync(appConfig);

            // Update in-memory cache.
            // this.userAppId = userAppId;
        }

        /// <inheritdoc/>
        public async Task DeleteUserAppIdAsync(string tenantId)
        {
            var appId = await this.GetUserAppIdAsync(tenantId);
            if (string.IsNullOrEmpty(appId))
            {
                // User App id isn't cached.
                return;
            }

            var appConfig = new AppConfigEntity()
            {
                PartitionKey = AppConfigTableName.SettingsUserAppIdPartition,
                RowKey = tenantId,
            };

            await this.repository.DeleteAsync(appConfig);

            // Clear in-memory cache.
            this.userAppId = null;
        }
    }
}

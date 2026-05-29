// <copyright file="UserDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>
namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;

    /// <summary>
    /// Repository of the user data stored in the table storage.
    /// </summary>
    public class UserDataRepository : BaseRepository<UserDataEntity>, IUserDataRepository
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UserDataRepository"/> class.
        /// </summary>
        /// <param name="logger">The logging service.</param>
        /// <param name="repositoryOptions">Options used to create the repository.</param>
        public UserDataRepository(
            ILogger<UserDataRepository> logger,
            IOptions<RepositoryOptions> repositoryOptions)
            : base(
                  logger,
                  storageAccountConnectionString: repositoryOptions.Value.StorageAccountConnectionString,
                  tableName: UserDataTableNames.TableName,
                  defaultPartitionKey: UserDataTableNames.UserDataPartition,
                  ensureTableExists: repositoryOptions.Value.EnsureTableExists)
        {
        }

        /// <inheritdoc/>
        public async Task<string> GetDeltaLinkAsync()
        {
            try
            {
                var operation = TableOperation.Retrieve<UsersSyncEntity>(UserDataTableNames.UsersSyncDataPartition, UserDataTableNames.AllUsersDeltaLinkRowKey);
                var result = await this.Table.ExecuteAsync(operation);
                var entity = result.Result as UsersSyncEntity;
                return entity?.Value;
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task SetDeltaLinkAsync(string deltaLink)
        {
            if (string.IsNullOrEmpty(deltaLink))
            {
                throw new ArgumentNullException(nameof(deltaLink));
            }

            var entity = new UsersSyncEntity()
            {
                PartitionKey = UserDataTableNames.UsersSyncDataPartition,
                RowKey = UserDataTableNames.AllUsersDeltaLinkRowKey,
                Value = deltaLink,
            };

            try
            {
                var operation = TableOperation.InsertOrReplace(entity);
                await this.Table.ExecuteAsync(operation);
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task<string> GetDeltaLinkAsync(string tenantId)
        {
            string deltaLink = string.Empty;
            try
            {
                string rowKey = $"{UserDataTableNames.AllUsersDeltaLinkRowKey}-{tenantId}";
                var operation = TableOperation.Retrieve<UsersSyncEntity>(UserDataTableNames.UsersSyncDataPartition, rowKey);
                var result = await this.Table.ExecuteAsync(operation);
                var entity = result.Result as UsersSyncEntity;
                deltaLink = entity?.Value;

                return deltaLink;
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task SetDeltaLinkAsync(string deltaLink, string tenantId)
        {
            if (string.IsNullOrEmpty(deltaLink))
            {
                throw new ArgumentNullException(nameof(deltaLink));
            }

            if (string.IsNullOrEmpty(tenantId))
            {
                throw new ArgumentNullException(nameof(tenantId));
            }

            string rowKey = $"{UserDataTableNames.AllUsersDeltaLinkRowKey}-{tenantId}";
            var entity = new UsersSyncEntity()
            {
                PartitionKey = UserDataTableNames.UsersSyncDataPartition,
                RowKey = rowKey,
                Value = deltaLink,
                TenantId = tenantId,
            };

            try
            {
                var operation = TableOperation.InsertOrReplace(entity);
                await this.Table.ExecuteAsync(operation);
            }
            catch (Exception ex)
            {
                this.Logger.LogError(ex, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// GetUserDataEntityByTenantIdAsync.
        /// </summary>
        /// <param name="tenantId">Tenant Id.</param>
        /// <returns>List of UserDataEntity.</returns>
        public async Task<IEnumerable<UserDataEntity>> GetUserDataEntityByTenantIdAsync(string tenantId)
        {
            string strFilter = TableQuery.GenerateFilterCondition("TenantId", QueryComparisons.Equal, tenantId);
            var result = await this.GetWithFilterAsync(strFilter, UserDataTableNames.UserDataPartition);

            return result;
        }
    }
}

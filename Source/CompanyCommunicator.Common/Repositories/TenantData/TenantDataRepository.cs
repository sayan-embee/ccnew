// <copyright file="TenantDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TenantData
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;

    /// <summary>
    /// Repository of the tenant data stored in the table storage.
    /// </summary>
    public class TenantDataRepository : BaseRepository<TenantDataEntity>, ITenantDataRepository
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TenantDataRepository"/> class.
        /// </summary>
        /// <param name="logger">The logging service.</param>
        /// <param name="repositoryOptions">Options used to create the repository.</param>
        public TenantDataRepository(
            ILogger<TenantDataRepository> logger,
            IOptions<RepositoryOptions> repositoryOptions)
            : base(
                  logger,
                  storageAccountConnectionString: repositoryOptions.Value.StorageAccountConnectionString,
                  tableName: TenantDataTableNames.TableName,
                  defaultPartitionKey: TenantDataTableNames.TenantDataPartition,
                  ensureTableExists: repositoryOptions.Value.EnsureTableExists)
        {
        }

        /// <inheritdoc/>
        public async Task EnsureTenantDataTableExistsAsync()
        {
            var exists = await this.Table.ExistsAsync();
            if (!exists)
            {
                await this.Table.CreateAsync();
            }
        }
    }
}

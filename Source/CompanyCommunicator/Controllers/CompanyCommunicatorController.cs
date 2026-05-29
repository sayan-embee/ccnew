// <copyright file="CompanyCommunicatorController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TenantData;

    /// <summary>
    /// Controller for the survey export data.
    /// </summary>
    [Route("api/companycommunicator")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    [ApiController]
    public class CompanyCommunicatorController : ControllerBase
    {
        private readonly ITenantDataRepository tenantDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="CompanyCommunicatorController"/> class.
        /// </summary>
        /// <param name="tenantDataRepository">Tenant data respository.</param>
        public CompanyCommunicatorController(
            ITenantDataRepository tenantDataRepository)
        {
            this.tenantDataRepository = tenantDataRepository ?? throw new ArgumentNullException(nameof(tenantDataRepository));
        }

        /// <summary>
        /// Get a sent notification by Id.
        /// </summary>
        /// <returns>Required sent notification.</returns>
        [HttpGet]
        [Route("tenantlist")]
        public async Task<IActionResult> GetSisterTenant()
        {
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json")
                .Build();
            var sistertenant = configuration.GetSection("SisterTenantId").Value.ToString();
            await Task.Delay(0);
            return this.Ok(sistertenant);
        }

        /// <summary>
        /// Get a list of all tenants.
        /// </summary>
        /// <returns>It returns list of tenant.</returns>
        [HttpGet]
        [Route("allTenantList")]
        public async Task<ActionResult<IEnumerable<TenantDataEntity>>> GetTenantsAsync()
        {
            var tenantEntities = await this.tenantDataRepository.GetAllAsync(
                TenantDataTableNames.TenantDataPartition);
            return this.Ok(tenantEntities);
        }
    }
}
// <copyright file="TeamDataController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Controller for the teams data.
    /// </summary>
    [Route("api/teamData")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class TeamDataController : ControllerBase
    {
        private readonly ITeamDataRepository teamDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamDataController"/> class.
        /// </summary>
        /// <param name="teamDataRepository">Team data repository instance.</param>
        public TeamDataController(ITeamDataRepository teamDataRepository)
        {
            this.teamDataRepository = teamDataRepository ?? throw new ArgumentNullException(nameof(teamDataRepository));
        }

        /// <summary>
        /// Get data for all teams.
        /// </summary>
        /// <param name="tenantId">Tenant Id.</param>
        /// <returns>A list of team data.</returns>
        [HttpGet]
        public async Task<IEnumerable<TeamData>> GetAllTeamDataAsync(string tenantId = "")
        {
            IEnumerable<TeamDataEntity> entities = null;
            if (string.IsNullOrEmpty(tenantId) || string.IsNullOrWhiteSpace(tenantId))
            {
                entities = await this.teamDataRepository.GetAllSortedAlphabeticallyByNameAsync();
            }
            else
            {
                entities = await this.teamDataRepository.GetAllSortedAlphabeticallyByNameAsync(tenantId);
            }

            var result = new List<TeamData>();
            foreach (var entity in entities)
            {
                var team = new TeamData
                {
                    Id = entity.TeamId,
                    Name = entity.Name,
                };
                result.Add(team);
            }

            return result;
        }

        /*
        /// <summary>
        /// Get data for all teams.
        /// </summary>
        /// <param name="tenantId">Tenant Id.</param>
        /// <returns>A list of team data.</returns>
        [HttpGet]
        [Route("getByTenantId")]
        public async Task<IEnumerable<TeamData>> GetAllTeamDataAsync(string tenantId)
        {
            var entities = await this.teamDataRepository.GetAllSortedAlphabeticallyByNameAsync(tenantId);
            var result = new List<TeamData>();
            foreach (var entity in entities)
            {
                var team = new TeamData
                {
                    Id = entity.TeamId,
                    Name = entity.Name,
                };
                result.Add(team);
            }

            return result;
       }*/
    }
}

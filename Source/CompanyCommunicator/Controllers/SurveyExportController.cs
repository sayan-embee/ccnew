// <copyright file="SurveyExportController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Helpers;

    /// <summary>
    /// Controller for the survey export data.
    /// </summary>
    [Route("api/surveyexport")]
    [ApiController]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class SurveyExportController : ControllerBase
    {
        private readonly CloudStorageHelper cloudStorageHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="SurveyExportController"/> class.
        /// </summary>
        /// <param name="cloudStorageHelper">Cloud storage helper service.</param>
        public SurveyExportController(CloudStorageHelper cloudStorageHelper)
        {
            this.cloudStorageHelper = cloudStorageHelper ?? throw new ArgumentNullException(nameof(cloudStorageHelper));
        }

        /// <summary>
        /// Get survey export data based on notification id.
        /// </summary>
        /// <param name="id">Notification Id.</param>
        /// <returns>List of survey details.</returns>
        [HttpPost]
        [Route("exportdata")]
        public async Task<IActionResult> GetSurveyExport(string id)
        {
            var result = await this.cloudStorageHelper.GetSurveryList(id);
            return this.Ok(result);
        }
    }
}
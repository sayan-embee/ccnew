// <copyright file="ReactionExportController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Helpers;
    using Newtonsoft.Json;

    /// <summary>
    /// Controller for the survey export data.
    /// </summary>
    [Route("api/reactionexport")]
    [ApiController]
    //[Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class ReactionExportController : ControllerBase
    {
        private readonly CloudStorageHelper cloudStorageHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="ReactionExportController"/> class.
        /// </summary>
        /// <param name="cloudStorageHelper">Cloud storage helper service.</param>
        public ReactionExportController(CloudStorageHelper cloudStorageHelper)
        {
            this.cloudStorageHelper = cloudStorageHelper ?? throw new ArgumentNullException(nameof(cloudStorageHelper));
        }

        /// <summary>
        /// Get reaction export data based on notification id.
        /// </summary>
        /// <param name="id">Notification Id.</param>
        /// <returns>List of reaction details.</returns>
        [HttpPost]
        [Route("exportdata")]
        public async Task<IActionResult> GetReactionExport(string id)
        {
            var result = await this.cloudStorageHelper.GetReactionList(id);
            return this.Ok(result);
        }

        /// <summary>
        /// Get reaction export data based on notification id.
        /// </summary>
        /// <param name="id">Notification Id.</param>
        /// <returns>List of reaction details.</returns>
        [HttpPost]
        [Route("exportReactionDetail")]
        public async Task<IActionResult> GetReactionDetailsExport(string id)
        {
            var result = await this.cloudStorageHelper.GetReactionListDetail(id);
            return this.Ok(result);
        }

        #region Get System Ip Address
        public class IPAddessModel
        {
            [JsonProperty("ipAddress")]
            public string IpAddress { get; set; }
        }
        [HttpGet("GetSystemIpAddress")]
        public async Task<ActionResult> GetSystemIpAddress()
        //public void GetSystemIpAddress()
        {
            try
            {
                IPAddessModel iPAddessModel = new IPAddessModel();
                string IPAddress = "";
                IPHostEntry Host = default(IPHostEntry);
                string Hostname = null;
                Hostname = System.Environment.MachineName;
                Host = Dns.GetHostEntry(Hostname);
                foreach (IPAddress IP in Host.AddressList)
                {
                    if (IP.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                    {
                        IPAddress = Convert.ToString(IP);
                        iPAddessModel.IpAddress = IPAddress;
                    }
                }
                Console.WriteLine(IPAddress);
                return this.Ok(iPAddessModel);
            }
            catch (Exception ex)
            {
                return this.Problem(ex.Message);
            }
        }
        #endregion
    }
}

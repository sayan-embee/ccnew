// <copyright file="TenantDataEntity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TenantData
{
    using Microsoft.Azure.Cosmos.Table;

    /// <summary>
    /// Teams data entity class.
    /// This entity holds the information about a team.
    /// </summary>
    public class TenantDataEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets the name of the tenamt.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether gets or sets IsPrimary for the tenamt.
        /// </summary>
        public bool IsPrimary { get; set; }
    }
}

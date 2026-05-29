// <copyright file="TenantDataTableNames.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TenantData
{
    /// <summary>
    /// Team data table names.
    /// </summary>
    public static class TenantDataTableNames
    {
        /// <summary>
        /// Table name for the team data table.
        /// </summary>
        public static readonly string TableName = "TenantData";

        /// <summary>
        /// Team data partition key name.
        /// </summary>
        public static readonly string TenantDataPartition = "TenantData";
    }
}

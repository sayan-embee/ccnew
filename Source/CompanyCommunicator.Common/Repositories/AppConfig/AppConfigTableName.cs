// <copyright file="AppConfigTableName.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories
{
    /// <summary>
    /// App config table information.
    /// </summary>
    public class AppConfigTableName
    {
        /// <summary>
        /// Table name for app config..
        /// </summary>
        public static readonly string TableName = "AppConfig";

        /// <summary>
        /// App settings partition.
        /// </summary>
        public static readonly string SettingsPartition = "Settings";

        /// <summary>
        /// Settings ServiceUrl partition.
        /// </summary>
        public static readonly string SettingsServiceUrlPartition = "SettingsServiceUrl";

        /// <summary>
        /// Settings UserAppId partition.
        /// </summary>
        public static readonly string SettingsUserAppIdPartition = "SettingsUserAppId";

        ///// <summary>
        ///// Service url row key.
        ///// </summary>
        // public static readonly string ServiceUrlRowKey = "ServiceUrl";

        ///// <summary>
        ///// User app id row key.
        ///// </summary>
        // public static readonly string UserAppIdRowKey = "UserAppId";
    }
}

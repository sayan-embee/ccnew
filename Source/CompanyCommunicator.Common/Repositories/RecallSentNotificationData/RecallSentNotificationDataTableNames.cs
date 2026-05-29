// <copyright file="RecallSentNotificationDataTableNames.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RecallSentNotificationData
{
    /// <summary>
    /// Sent notification data table names.
    /// </summary>
    public static class RecallSentNotificationDataTableNames
    {
        /// <summary>
        /// Table name for the sent notification data table.
        /// </summary>
        public static readonly string TableName = "RecallSentNotificationData";

        /// <summary>
        /// Default partition - should not be used.
        /// </summary>
        public static readonly string DefaultPartition = "Default";
    }
}

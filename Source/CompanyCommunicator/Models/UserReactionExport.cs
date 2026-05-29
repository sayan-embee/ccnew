// <copyright file="UserReactionExport.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;

    /// <summary>
    /// UserReactionExport.
    /// </summary>
    public class UserReactionExport
    {
        /// <summary>
        /// Gets or sets a LikeCount.
        /// </summary>
        public int LikeCount { get; set; }

        /// <summary>
        /// Gets or sets a HeartCount.
        /// </summary>
        public int HeartCount { get; set; }

        /// <summary>
        /// Gets or sets a LaughCount.
        /// </summary>
        public int LaughCount { get; set; }

        /// <summary>
        /// Gets or sets a SurprisedCount.
        /// </summary>
        public int SurprisedCount { get; set; }

        /// <summary>
        /// Gets or sets a SadCount.
        /// </summary>
        public int SadCount { get; set; }

        /// <summary>
        /// Gets or sets a AngryCount.
        /// </summary>
        public int AngryCount { get; set; }
    }
}

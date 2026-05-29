// <copyright file="UserReaction.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Models
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;

    /// <summary>
    /// Base QuestionAnswer model class.
    /// </summary>
    public class UserReaction
    {
        /// <summary>
        /// Gets or sets the FromId.
        /// </summary>
        public string FromId { get; set; }

        /// <summary>
        /// Gets or sets the Name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the AadId.
        /// </summary>
        public string AadId { get; set; }

        /// <summary>
        /// Gets or sets the ActivityId.
        /// </summary>
        public string ActivityId { get; set; }

        /// <summary>
        /// Gets or sets the ReactionType.
        /// </summary>
        public string ReactionType { get; set; }

        /// <summary>
        /// Gets or sets the Email.
        /// </summary>
        public string Email { get; set; }

        /// <summary>
        /// Gets or sets the Name.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets the Name.
        /// </summary>
        public string NotificationId { get; set; }

        /// <summary>
        /// Gets or sets the ReactionDate.
        /// </summary>
        public DateTime? ReactionDate { get; set; }
    }
}

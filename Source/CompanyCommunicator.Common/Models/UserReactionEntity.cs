// <copyright file="UserReactionEntity.cs" company="Microsoft">
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
    /// UserReactionEntity.
    /// </summary>
    public class UserReactionEntity : TableEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UserReactionEntity"/> class.
        /// </summary>
        public UserReactionEntity()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UserReactionEntity"/> class.
        /// </summary>
        /// <param name="activityId">Activity Id.</param>
        /// <param name="fromId">From Id.</param>
        public UserReactionEntity(string activityId, string fromId)
        {
            this.PartitionKey = activityId;
            this.RowKey = fromId;
        }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string FromId { get; set; }

        /// <summary>
        /// Gets or sets a Name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets a AadId.
        /// </summary>
        public string AadId { get; set; }

        /// <summary>
        /// Gets or sets a ActivityId.
        /// </summary>
        public string ActivityId { get; set; }

        /// <summary>
        /// Gets or sets a ReactionType.
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

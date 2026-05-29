// <copyright file="ReportSentNotificationDetailModel.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Models
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;

    /// <summary>
    /// Report Sent Notification Detail Model.
    /// </summary>
    public class ReportSentNotificationDetailModel
    {
        /// <summary>
        /// Gets or sets Message Title.
        /// </summary>
        public string MessageTitle { get; set; }

        /// <summary>
        /// Gets or sets notification Id.
        /// </summary>
        public string NotificationId { get; set; }

        /// <summary>
        /// Gets or sets activity Id.
        /// </summary>
        public string ActivityId { get; set; }

        /// <summary>
        /// Gets or sets message sent on datetime.
        /// </summary>
        public DateTime? SentDate { get; set; }

        /// <summary>
        /// Gets or sets tenant Id.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets Delivery Status.
        /// </summary>
        public string DeliveryStatus { get; set; }

        /// <summary>
        /// Gets or sets Name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets Email.
        /// </summary>
        public string Email { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether gets or sets ReadReceipt.
        /// </summary>
        public bool ReadReceipt { get; set; }

        /// <summary>
        /// Gets or sets ReadReceiptDate datetime.
        /// </summary>
        public DateTime? ReadReceiptDate { get; set; }

        /// <summary>
        /// Gets or sets Email.
        /// </summary>
        public string UserId { get; set; }

        /// <summary>
        /// Gets or sets Reaction.
        /// </summary>
        public string Reaction { get; set; }

        /// <summary>
        /// Gets or sets ReactionDate datetime.
        /// </summary>
        public DateTime? ReactionDate { get; set; }

        /// <summary>
        /// Gets or sets TenantName.
        /// </summary>
        public string TenantName { get; set; }
    }
}

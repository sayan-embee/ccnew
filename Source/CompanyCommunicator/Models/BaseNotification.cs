// <copyright file="BaseNotification.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    using System;

    /// <summary>
    /// Base notification model class.
    /// </summary>
    public class BaseNotification
    {
        /// <summary>
        /// Gets or sets Id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets Title value.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the Image Link value.
        /// </summary>
        public string ImageLink { get; set; }

        /// <summary>
        /// Gets or sets the Summary value.
        /// </summary>
        public string Summary { get; set; }

        /// <summary>
        /// Gets or sets the Author value.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets the Button Title value.
        /// </summary>
        public string ButtonTitle { get; set; }

        /// <summary>
        /// Gets or sets the Button Link value.
        /// </summary>
        public string ButtonLink { get; set; }

        /// <summary>
        /// Gets or sets the Buttons value.
        /// </summary>
        public string Buttons { get; set; }

        /// <summary>
        /// Gets or sets the TemplateType value.
        /// </summary>
        public string TemplateType { get; set; }

        /// <summary>
        /// Gets or sets the TenantId value.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets the Send Type value.
        /// </summary>
        public string SendTypeId { get; set; }

        /// <summary>
        /// Gets or sets the AdaptivecardContent.
        /// </summary>
        public string AdaptiveCardContent { get; set; }

        /// <summary>
        /// Gets or sets the EmailBody.
        /// </summary>
        public string EmailBody { get; set; }

        /// <summary>
        /// Gets or sets the EmailBody.
        /// </summary>
        public string EmailTitle { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether gets or sets the IsScheduled value.
        /// </summary>
        public bool IsScheduled { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether gets or sets the IsImportant value.
        /// </summary>
        public bool IsImportant { get; set; }

        /// <summary>
        /// Gets or sets the Created DateTime value.
        /// </summary>
        public DateTime CreatedDateTime { get; set; }

        /// <summary>
        /// Gets or sets the CsvLink value.
        /// </summary>
        public string CsvLink { get; set; }

        /// <summary>
        /// Gets or sets the AdditionalFileLink  value.
        /// </summary>
        public string AdditionalFileLink { get; set; }

        /// <summary>
        /// Gets or sets the TenantName value.
        /// </summary>
        public string TenantName { get; set; }

        /// <summary>
        /// Gets or sets the AuthorTeamId value.
        /// </summary>
        public string AuthorTeamId { get; set; }

        /// <summary>
        /// Gets or sets the AuthorTeamName value.
        /// </summary>
        public string AuthorTeamName { get; set; }

        /// <summary>
        /// Gets or sets the AuthorChannelId value.
        /// </summary>
        public string AuthorChannelId { get; set; }

        /// <summary>
        /// Gets or sets the AuthorChannelName value.
        /// </summary>
        public string AuthorChannelName { get; set; }
    }
}

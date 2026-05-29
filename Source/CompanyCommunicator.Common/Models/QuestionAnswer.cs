// <copyright file="QuestionAnswer.cs" company="Microsoft">
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
    using Newtonsoft.Json;

    /// <summary>
    /// Base QuestionAnswer model class.
    /// </summary>
    public class QuestionAnswer
    {
        /// <summary>
        /// Gets or sets FromId.
        /// </summary>
        public string FromId { get; set; }

        /// <summary>
        /// Gets or sets Name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets Email.
        /// </summary>
        public string Email { get; set; }

        /// <summary>
        /// Gets or sets TenantId.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets AadId.
        /// </summary>
        public string AadId { get; set; }

        /// <summary>
        /// Gets or sets NotificationId.
        /// </summary>
        public string NotificationId { get; set; }

        /// <summary>
        /// Gets or sets Questions.
        /// </summary>
        public string Questions { get; set; }

        /// <summary>
        /// Gets or sets Title.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets Author.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets answer0.
        /// </summary>
        [JsonProperty("answer0")]
        public string Answer0 { get; set; }

        /// <summary>
        /// Gets or sets answer1.
        /// </summary>
        [JsonProperty("answer1")]
        public string Answer1 { get; set; }

        /// <summary>
        /// Gets or sets answer2.
        /// </summary>
        [JsonProperty("answer2")]
        public string Answer2 { get; set; }

        /// <summary>
        /// Gets or sets answer3.
        /// </summary>
        [JsonProperty("answer3")]
        public string Answer3 { get; set; }

        /// <summary>
        /// Gets or sets answer4.
        /// </summary>
        [JsonProperty("answer4")]
        public string Answer4 { get; set; }

        /// <summary>
        /// Gets or sets answer5.
        /// </summary>
        [JsonProperty("answer5")]
        public string Answer5 { get; set; }

        /// <summary>
        /// Gets or sets answer5.
        /// </summary>
        [JsonProperty("answer6")]
        public string Answer6 { get; set; }

        /// <summary>
        /// Gets or sets answer5.
        /// </summary>
        [JsonProperty("answer7")]
        public string Answer7 { get; set; }

        /// <summary>
        /// Gets or sets a SubmittedOn.
        /// </summary>
        [JsonProperty("SubmittedOn")]
        public DateTime? SubmittedOn { get; set; }

        /// <summary>
        /// Gets or sets ConversationId.
        /// </summary>
        public string ConversationId { get; set; }

        /// <summary>
        /// Gets or sets ActivityId.
        /// </summary>
        public string ActivityId { get; set; }
    }
}

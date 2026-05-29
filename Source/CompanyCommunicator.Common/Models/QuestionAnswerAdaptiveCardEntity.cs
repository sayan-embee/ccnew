// <copyright file="QuestionAnswerAdaptiveCardEntity.cs" company="Microsoft">
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
    /// Gets or sets QuestionAnswerAdaptiveCardEntity.
    /// </summary>
    public class QuestionAnswerAdaptiveCardEntity : TableEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="QuestionAnswerAdaptiveCardEntity"/> class.
        /// </summary>
        public QuestionAnswerAdaptiveCardEntity()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="QuestionAnswerAdaptiveCardEntity"/> class.
        /// </summary>
        /// <param name="notificationId">Notification Id.</param>
        /// <param name="notification">Notification.</param>
        public QuestionAnswerAdaptiveCardEntity(string notificationId, string notification)
        {
            this.PartitionKey = notificationId;
            this.RowKey = notification;
        }

        /// <summary>
        /// Gets or sets a Title.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets a Author.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets a NotificationId.
        /// </summary>
        public string NotificationId { get; set; }

        /// <summary>
        /// Gets or sets a Questions.
        /// </summary>
        public string Questions { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string FromId { get; set; }

        /// <summary>
        /// Gets or sets a Name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets a Email.
        /// </summary>
        public string Email { get; set; }

        /// <summary>
        /// Gets or sets a AadId.
        /// </summary>
        public string AadId { get; set; }

        /// <summary>
        /// Gets or sets a Question0.
        /// </summary>
        public string Question0 { get; set; }

        /// <summary>
        /// Gets or sets a Question1.
        /// </summary>
        public string Question1 { get; set; }

        /// <summary>
        /// Gets or sets a Question2.
        /// </summary>
        public string Question2 { get; set; }

        /// <summary>
        /// Gets or sets a Question3.
        /// </summary>
        public string Question3 { get; set; }

        /// <summary>
        /// Gets or sets a Question4.
        /// </summary>
        public string Question4 { get; set; }

        /// <summary>
        /// Gets or sets a Question5.
        /// </summary>
        public string Question5 { get; set; }

        /// <summary>
        /// Gets or sets a Question6.
        /// </summary>
        public string Question6 { get; set; }

        /// <summary>
        /// Gets or sets a Question7.
        /// </summary>
        public string Question7 { get; set; }

        /// <summary>
        /// Gets or sets a Answer0.
        /// </summary>
        public string Answer0 { get; set; }

        /// <summary>
        /// Gets or sets a Answer1.
        /// </summary>
        public string Answer1 { get; set; }

        /// <summary>
        /// Gets or sets a Answer2.
        /// </summary>
        public string Answer2 { get; set; }

        /// <summary>
        /// Gets or sets a Answer3.
        /// </summary>
        public string Answer3 { get; set; }

        /// <summary>
        /// Gets or sets a Answer4.
        /// </summary>
        public string Answer4 { get; set; }

        /// <summary>
        /// Gets or sets a Answer5.
        /// </summary>
        public string Answer5 { get; set; }

        /// <summary>
        /// Gets or sets a Answer6.
        /// </summary>
        public string Answer6 { get; set; }

        /// <summary>
        /// Gets or sets a Answer7.
        /// </summary>
        public string Answer7 { get; set; }

        /// <summary>
        /// Gets or sets a SubmittedOn.
        /// </summary>
        public DateTime? SubmittedOn { get; set; }

        /// <summary>
        /// Gets or sets a TenantId.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets a ConversationId.
        /// </summary>
        public string ConversationId { get; set; }

        /// <summary>
        /// Gets or sets a ActivityId.
        /// </summary>
        public string ActivityId { get; set; }
    }
}

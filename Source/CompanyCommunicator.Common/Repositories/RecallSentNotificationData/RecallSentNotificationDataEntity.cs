// <copyright file="RecallSentNotificationDataEntity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RecallSentNotificationData
{
    using System;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;

    /// <summary>
    /// Sent notification entity class.
    /// This entity holds all of the information about a recipient and the results for
    /// a notification having been sent to that recipient.
    /// </summary>
    public class RecallSentNotificationDataEntity : SentNotificationDataEntity
    {
    }
}

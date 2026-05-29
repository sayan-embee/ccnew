// <copyright file="IUserDataService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.User
{
    using System.Threading.Tasks;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;

    /// <summary>
    /// User Data Service.
    /// </summary>
    public interface IUserDataService
    {
        /// <summary>
        /// Add user data in Table Storage.
        /// </summary>
        /// <param name="activity">Bot conversation update activity instance.</param>
        /// <param name="teamsChannelAccount">teamsChannelAccount.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public Task SaveUserDataAsync(IConversationUpdateActivity activity, TeamsChannelAccount teamsChannelAccount = null);

        /// <summary>
        /// Add author data in Table Storage.
        /// </summary>
        /// <param name="activity">Bot conversation update activity instance.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public Task SaveAuthorDataAsync(IConversationUpdateActivity activity);

        /// <summary>
        /// Remove personal data in table storage.
        /// </summary>
        /// <param name="activity">Bot conversation update activity instance.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public Task RemoveUserDataAsync(IConversationUpdateActivity activity);

        /// <summary>
        /// Remove author data in table storage.
        /// </summary>
        /// <param name="activity">Bot conversation update activity instance.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public Task RemoveAuthorDataAsync(IConversationUpdateActivity activity);
    }
}

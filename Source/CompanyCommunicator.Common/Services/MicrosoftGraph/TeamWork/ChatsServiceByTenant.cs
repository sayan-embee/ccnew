// <copyright file="ChatsServiceByTenant.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Graph;

    /// <summary>
    /// Chats Service.
    /// </summary>
    internal class ChatsServiceByTenant : IChatsServiceByTenant
    {
        private readonly IGraphServiceClient graphServiceClient;
        private readonly IAppManagerServiceByTenant appManagerServiceByTenant;

        /// <summary>
        /// Initializes a new instance of the <see cref="ChatsServiceByTenant"/> class.
        /// </summary>
        /// <param name="graphServiceClient">Graph service client.</param>
        /// <param name="appManagerServiceByTenant">App manager service.</param>
        internal ChatsServiceByTenant(
            IGraphServiceClient graphServiceClient,
            IAppManagerServiceByTenant appManagerServiceByTenant)
        {
            this.graphServiceClient = graphServiceClient ?? throw new ArgumentNullException(nameof(graphServiceClient));
            this.appManagerServiceByTenant = appManagerServiceByTenant ?? throw new ArgumentNullException(nameof(appManagerServiceByTenant));
        }

        /// <inheritdoc/>
        public async Task<string> GetChatThreadIdAsync(string userId, string appId)
        {
            if (string.IsNullOrWhiteSpace(userId))
            {
                throw new ArgumentNullException(nameof(userId));
            }

            if (string.IsNullOrWhiteSpace(appId))
            {
                throw new ArgumentNullException(nameof(appId));
            }

            var installationId = await this.appManagerServiceByTenant.GetAppInstallationIdForUserAsync(appId, userId);
            var chat = await this.graphServiceClient.Users[userId]
                .Teamwork
                .InstalledApps[installationId]
                .Chat
                .Request()
                .WithMaxRetry(GraphConstants.MaxRetry)
                .GetAsync();

            return chat?.Id;
        }
    }
}

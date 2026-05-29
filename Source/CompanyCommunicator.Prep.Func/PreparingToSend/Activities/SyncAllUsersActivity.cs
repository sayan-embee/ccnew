// <copyright file="SyncAllUsersActivity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Functions.Worker;
    using Microsoft.DurableTask;
    using Microsoft.DurableTask.Client;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Resources;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend.Extensions;

    /// <summary>
    /// Syncs all users to Sent notification table.
    /// </summary>
    public class SyncAllUsersActivity
    {
        private readonly IUserDataRepository userDataRepository;
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly IUsersService usersService;
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly IStringLocalizer<Strings> localizer;
        private readonly IGraphServiceFactoryByTenant graphServiceFactoryByTenant;

        /// <summary>
        /// Initializes a new instance of the <see cref="SyncAllUsersActivity"/> class.
        /// </summary>
        /// <param name="userDataRepository">User Data repository.</param>
        /// <param name="sentNotificationDataRepository">Sent notification data repository.</param>
        /// <param name="usersService">Users service.</param>
        /// <param name="notificationDataRepository">Notification data entity repository.</param>
        /// <param name="localizer">Localization service.</param>
        /// <param name="graphServiceFactoryByTenant">graphServiceFactoryByTenant.</param>
        public SyncAllUsersActivity(
            IUserDataRepository userDataRepository,
            ISentNotificationDataRepository sentNotificationDataRepository,
            IUsersService usersService,
            INotificationDataRepository notificationDataRepository,
            IStringLocalizer<Strings> localizer,
            IGraphServiceFactoryByTenant graphServiceFactoryByTenant)
        {
            this.userDataRepository = userDataRepository ?? throw new ArgumentNullException(nameof(userDataRepository));
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentNullException(nameof(sentNotificationDataRepository));
            this.usersService = usersService ?? throw new ArgumentNullException(nameof(usersService));
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.localizer = localizer ?? throw new ArgumentNullException(nameof(localizer));
            this.graphServiceFactoryByTenant = graphServiceFactoryByTenant ?? throw new ArgumentNullException(nameof(graphServiceFactoryByTenant));
        }

        /// <summary>
        /// Syncs all users to Sent notification table.
        /// </summary>
        /// <param name="notification">Notification.</param>
        /// <param name="log">Logging service.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [Function(FunctionNames.SyncAllUsersActivity)]
        public async Task RunAsync([ActivityTrigger] NotificationDataEntity notification, ILogger log)
        {
            if (notification == null)
            {
                throw new ArgumentNullException(nameof(notification));
            }

            if (log == null)
            {
                throw new ArgumentNullException(nameof(log));
            }

            // Sync all users.
            if (notification.SendTypeId != "5")
            {
                await this.SyncAllUsers(notification.Id, notification, log);
            }

            // Get users.
            var users = await this.userDataRepository.GetUserDataEntityByTenantIdAsync(notification.TenantId);

            // users = users.Where(u => !u.TenantId.IsNullOrEmpty() && u.TenantId.ToString() == notification.TenantId.ToString());

            // Store in sent notification table.
            var recipients = users.Select(
                user => user.CreateInitialSentNotificationDataEntity(partitionKey: notification.Id));
            //log.LogInformation($"SyncAllUsers for notification id {notification.Id} and tenant id {notification.TenantId} user count is {users.Count()} and recipients count is {recipients.Count()}");
            await this.sentNotificationDataRepository.BatchInsertOrMergeAsync(recipients);
        }

        /// <summary>
        /// Syncs delta changes only.
        /// </summary>
        private async Task SyncAllUsers(string notificationId, NotificationDataEntity notification, ILogger log)
        {
            // Sync users
            log.LogInformation($"SyncAllUsers for notification id {notificationId} and tenant id {notification.TenantId}");
            var deltaLink = string.Empty;
            if (!string.IsNullOrEmpty(notification.TenantId))
            {
                deltaLink = await this.userDataRepository.GetDeltaLinkAsync(notification.TenantId);
                log.LogInformation($"deltaLink for notification id {notificationId} and tenant id {notification.TenantId}");
            }
            else
            {
                deltaLink = await this.userDataRepository.GetDeltaLinkAsync();
                log.LogInformation($"deltaLink for notification id {notificationId}");
            }

            (IEnumerable<User>, string) tuple = (new List<User>(), string.Empty);
            try
            {
                if (!string.IsNullOrEmpty(deltaLink))
                {
                    if (string.IsNullOrEmpty(notification.TenantId))
                    {
                        log.LogInformation($"GetAllUsersAsync deltaLink : {deltaLink}");
                        tuple = await this.usersService.GetAllUsersAsync(deltaLink);
                    }
                    else
                    {
                        log.LogInformation($"GetAllUsersAsync deltaLink : {deltaLink} and tenant id {notification.TenantId}");
                        var usersServiceByTenant = this.graphServiceFactoryByTenant.GetUsersService(notification.TenantId);
                        tuple = await usersServiceByTenant.GetAllUsersAsync(deltaLink);
                    }
                }
            }
            catch (ServiceException exception)
            {
                var errorMessage = this.localizer.GetString("FailedToGetAllUsersFormat", exception.StatusCode, exception.Message);
                await this.notificationDataRepository.SaveWarningInNotificationDataEntityAsync(notificationId, errorMessage);
                return;
            }

            // process users.
            var users = tuple.Item1;
            if (!users.IsNullOrEmpty())
            {
                var maxParallelism = Math.Min(users.Count(), 30);

                IUsersServiceByTenant usersServiceByTenant = null;
                if (!string.IsNullOrEmpty(notification.TenantId))
                {
                    usersServiceByTenant = this.graphServiceFactoryByTenant.GetUsersService(notification.TenantId);
                }

                log.LogInformation($"ProcessUserAsync for tuple : {tuple}");

                // await users.ForEachAsync(maxParallelism, this.ProcessUserAsync);
                await this.ProcessUserAsync(notification, users, maxParallelism, usersServiceByTenant);
            }

            // Store delta link
            if (!string.IsNullOrEmpty(tuple.Item2))
            {
                if (!string.IsNullOrEmpty(notification.TenantId))
                {
                    await this.userDataRepository.SetDeltaLinkAsync(tuple.Item2, notification.TenantId);
                }
                else
                {
                    await this.userDataRepository.SetDeltaLinkAsync(tuple.Item2);
                }
            }
        }

        private async Task ProcessUserAsync(NotificationDataEntity notification, IEnumerable<User> users, int maxParallelism, IUsersServiceByTenant userServiceByTenant)
        {
            await users.ForEachAsync(
                maxParallelism,
                async user =>
                {
                        // Delete users who were removed.
                        if (user.AdditionalData?.ContainsKey("@removed") == true)
                    {
                        var localUser = await this.userDataRepository.GetAsync(UserDataTableNames.UserDataPartition, user.Id);
                        if (localUser != null)
                        {
                            await this.userDataRepository.DeleteAsync(localUser);
                        }

                        return;
                    }

                        // skip Guest users.
                        if (string.Equals(user.UserType, "Guest", StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    // skip users who do not have teams license.
                        try
                    {
                        var hasTeamsLicense = true;

                        if (userServiceByTenant != null)
                        {
                            hasTeamsLicense = await userServiceByTenant.HasTeamsLicenseAsync(user.Id);
                        }
                        else
                        {
                            hasTeamsLicense = await this.usersService.HasTeamsLicenseAsync(user.Id);
                        }

                        if (!hasTeamsLicense)
                        {
                            return;
                        }
                    }
                    catch (ServiceException)
                    {
                        // Failed to get user's license details. Will skip the user.
                        return;
                    }

                        // Store user.
                        await this.userDataRepository.InsertOrMergeAsync(
                        new UserDataEntity()
                        {
                            PartitionKey = UserDataTableNames.UserDataPartition,
                            RowKey = user.Id,
                            AadId = user.Id,
                            Email = user.Mail,
                            Name = user.DisplayName,
                            Upn = user.UserPrincipalName,
                            TenantId = notification.TenantId,
                        });
                });
        }

        private async Task ProcessUserAsync(User user)
        {
            // Delete users who were removed.
            if (user.AdditionalData?.ContainsKey("@removed") == true)
            {
                var localUser = await this.userDataRepository.GetAsync(UserDataTableNames.UserDataPartition, user.Id);
                if (localUser != null)
                {
                    await this.userDataRepository.DeleteAsync(localUser);
                }

                return;
            }

            // skip Guest users.
            if (string.Equals(user.UserType, "Guest", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            // skip users who do not have teams license.
            try
            {
                var hasTeamsLicense = await this.usersService.HasTeamsLicenseAsync(user.Id);
                if (!hasTeamsLicense)
                {
                    return;
                }
            }
            catch (ServiceException)
            {
                // Failed to get user's license details. Will skip the user.
                return;
            }

            // Store user.
            await this.userDataRepository.InsertOrMergeAsync(
                new UserDataEntity()
                {
                    PartitionKey = UserDataTableNames.UserDataPartition,
                    RowKey = user.Id,
                    AadId = user.Id,
                    Email = user.Mail,
                    Name = user.DisplayName,
                    Upn = user.UserPrincipalName,
                });
        }
    }
}

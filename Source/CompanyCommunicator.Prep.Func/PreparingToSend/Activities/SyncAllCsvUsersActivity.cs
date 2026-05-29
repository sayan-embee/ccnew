// <copyright file="SyncAllCsvUsersActivity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
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
    using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.Export.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend.Extensions;

    /// <summary>
    /// Syncs all users to Sent notification table.
    /// </summary>
    public class SyncAllCsvUsersActivity
    {
        private readonly IUserDataRepository userDataRepository;
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly IUsersService usersService;
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly IStringLocalizer<Strings> localizer;
        private readonly IGraphServiceFactoryByTenant graphServiceFactoryByTenant;
               /// <summary>
        /// Initializes a new instance of the <see cref="SyncAllCsvUsersActivity"/> class.
        /// </summary>
        /// <param name="userDataRepository">User Data repository.</param>
        /// <param name="sentNotificationDataRepository">Sent notification data repository.</param>
        /// <param name="usersService">Users service.</param>
        /// <param name="notificationDataRepository">Notification data entity repository.</param>
        /// <param name="localizer">Localization service.</param>
        /// <param name="graphServiceFactoryByTenant">graphServiceFactoryByTenant.</param>
        /// <param name="log">The logger factory.</param>
        public SyncAllCsvUsersActivity(
            IUserDataRepository userDataRepository,
            ISentNotificationDataRepository sentNotificationDataRepository,
            IUsersService usersService,
            INotificationDataRepository notificationDataRepository,
            IStringLocalizer<Strings> localizer,
            IGraphServiceFactoryByTenant graphServiceFactoryByTenant
            )
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
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [Function(FunctionNames.SyncAllCsvUsersActivity)]
        public async Task RunAsync([ActivityTrigger] NotificationDataEntity notification, ILogger log)
        {
            if (notification == null)
            {
                throw new ArgumentNullException(nameof(notification));
            }

            // Sync all csv users.
            var adUsers = await this.SyncAllCsvUsers(notification, log);

            if (adUsers.Any())
            {
                // Get users.
                string strFilter = TableQuery.GenerateFilterCondition("TenantId", QueryComparisons.Equal, notification.TenantId);
                var users = await this.userDataRepository.GetWithFilterAsync(strFilter, UserDataTableNames.UserDataPartition);


                // var users = await this.userDataRepository.GetAllAsync(partition: "UserData");
                if (users.Any())
                {

                    // users = users.Where(u => !u.TenantId.IsNullOrEmpty() && u.TenantId.ToString() == notification.TenantId.ToString());
                    IList<UserDataEntity> csvUsersOnly = new List<UserDataEntity>();
                    if (users.Any())
                    {
                        var results = (from u in users
                                       join ad in adUsers
                                       on u.AadId equals ad.Id
                                       select u as UserDataEntity);
                        if (results.Any())
                        {
                            csvUsersOnly = results.ToList();
                        }

                        //foreach (var item in adUsers)
                        //{
                        //    var userItemList = users.Where(x => x.AadId.ToString().ToLower() == item.Id.ToString().ToLower()).ToList();
                        //    if (userItemList.Any())
                        //    {
                        //        csvUsersOnly.Add(userItemList[0]);
                        //    }
                        //}
                    }

                    if (csvUsersOnly.Any())
                    {
                        // Store in sent notification table.
                        var recipients = csvUsersOnly.Select(
                            user => user.CreateInitialSentNotificationDataEntity(partitionKey: notification.Id));
                        log.LogInformation($"SyncAllCSVUsers for notification id {notification.Id} and tenant id {notification.TenantId} user count is {users.Count()} and recipients count is {recipients.Count()}");
                        await this.sentNotificationDataRepository.BatchInsertOrMergeAsync(recipients);
                    }
                }
            }
        }

        /// <summary>
        /// Syncs delta changes only.
        /// </summary>
        private async Task<IEnumerable<User>> SyncAllCsvUsers(NotificationDataEntity notification, ILogger log)
        {
            IEnumerable<string> emailIds = new List<string>();
            IEnumerable<User> users = new List<User>();

            try
            {
                emailIds = this.GetUserEmailsFromCsv(notification.CsvLink);
            }
            catch (Exception exception)
            {
                var errorMessage = this.localizer.GetString("FailedToGetAllUsersFormat", exception.Message);
                await this.notificationDataRepository.SaveWarningInNotificationDataEntityAsync(notification.Id, errorMessage);
                return users;
            }

            if (emailIds.Any())
            {
                try
                {
                    var usersService = this.graphServiceFactoryByTenant.GetUsersService(notification.TenantId);
                    users = await usersService.GetBatchByUserIds(emailIds.ToList().AsGroups(), "EMAIL");
                    // users = await this.usersService.GetBatchByUserIds(emailIds.ToList().AsGroups(), "EMAIL");
                }
                catch (ServiceException exception)
                {
                    var errorMessage = this.localizer.GetString("FailedToGetAllUsersFormat", exception.StatusCode, exception.Message);
                    await this.notificationDataRepository.SaveWarningInNotificationDataEntityAsync(notification.Id, errorMessage);
                    return users;
                }

                // process users.
                if (!users.IsNullOrEmpty() && users.Any())
                {
                    var maxParallelism = Math.Min(users.Count(), 30);

                    IUsersServiceByTenant usersServiceByTenant = null;
                    if (!string.IsNullOrEmpty(notification.TenantId))
                    {
                        usersServiceByTenant = this.graphServiceFactoryByTenant.GetUsersService(notification.TenantId);
                    }

                    // await users.ForEachAsync(maxParallelism, this.ProcessUserAsync);
                    await this.ProcessUserAsync(notification, users, maxParallelism, usersServiceByTenant);
                }
            }

            return users;
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
                });
        }

        /// <summary>
        /// Read the csv file from url.
        /// </summary>
        /// <param name="url">CSV file url.</param>
        /// <returns>List of users.</returns>
        private IEnumerable<string> GetUserEmailsFromCsv(string url)
        {

            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
            string line;
            StreamReader sr = new StreamReader(resp.GetResponseStream());
            List<string> splitted = new List<string>();
            string[] row;
            int index = 0;
            while ((line = sr.ReadLine()) != null)
            {
                if (index > 0)
                {
                    row = line.Split(',');
                    if (!string.IsNullOrWhiteSpace(row[0]))
                    {
                        if (row[0].Contains("'"))
                        {
                            splitted.Add(row[0].Replace("'", "''"));
                        }
                        else
                        {
                            splitted.Add(row[0]);
                        }
                        //splitted.Add(row[0]);
                    }
                }

                index++;
            }

            return splitted;
        }
    }
}

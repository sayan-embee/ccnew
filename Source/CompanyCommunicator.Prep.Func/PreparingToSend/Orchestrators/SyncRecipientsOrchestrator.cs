// <copyright file="SyncRecipientsOrchestrator.cs" company="Microsoft">
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
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;

    /// <summary>
    /// Syncs target set of recipients to Sent notification table.
    /// </summary>
    public static class SyncRecipientsOrchestrator
    {
        /// <summary>
        /// Fetch recipients and store them in Azure storage.
        /// </summary>
        /// <param name="context">Durable orchestration context.</param>
        /// <param name="log">Logging service.</param>
        /// <returns><see cref="Task"/> representing the asynchronous operation.</returns>
        [Function(FunctionNames.SyncRecipientsOrchestrator)]
        public static async Task RunOrchestrator(
            [OrchestrationTrigger] TaskOrchestrationContext context)
        {
            var log = context.CreateReplaySafeLogger(nameof(Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend.FunctionNames));
            var notification = context.GetInput<NotificationDataEntity>();

            // Update notification status.
            await context.CallActivityAsync(FunctionNames.UpdateNotificationStatusActivity, (notification.Id, NotificationStatus.SyncingRecipients), FunctionSettings.DefaultRetryOptions);

            // All users.
            if (notification.AllUsers)
            {
                await context.CallActivityAsync(FunctionNames.SyncAllUsersActivity, notification, FunctionSettings.DefaultRetryOptions);
                return;
            }

            // Members of specific teams.
            if (notification.Rosters.Any())
            {
                var tasks = new List<Task>();
                foreach (var teamId in notification.Rosters)
                {
                    var task = context.CallActivityAsync(FunctionNames.SyncTeamMembersActivity, (notification.Id, teamId), FunctionSettings.DefaultRetryOptions);
                    tasks.Add(task);
                }

                // Fan-Out Fan-In.
                await Task.WhenAll(tasks);
                return;
            }

            // Members of M365 groups, DG or SG.
            if (notification.Groups.Any())
            {
                var tasks = new List<Task>();
                foreach (var groupId in notification.Groups)
                {
                    var task = context.CallActivityAsync(FunctionNames.SyncGroupMembersActivity, (notification.Id, groupId, notification.TenantId), FunctionSettings.DefaultRetryOptions);

                    tasks.Add(task);
                }

                // Fan-Out Fan-In
                await Task.WhenAll(tasks);
                return;
            }

            // General channel of teams.
            if (notification.Teams.Any())
            {
                await context.CallActivityAsync(FunctionNames.SyncTeamsActivity, notification, FunctionSettings.DefaultRetryOptions);
                return;
            }

            // Only CSV Users.
            log.LogInformation($"Syncing csv users for notification id : {notification.Id} file url : {notification.CsvLink}");
            if (!string.IsNullOrEmpty(notification.CsvLink))
            {
                await context.CallActivityAsync(FunctionNames.SyncAllCsvUsersActivity, notification, FunctionSettings.DefaultRetryOptions);
                return;
            }

            // Invalid audience.
            var errorMessage = $"Invalid audience select for notification id: {notification.Id}";
            log.LogError(errorMessage);
            throw new ArgumentException(errorMessage);
        }
    }
}
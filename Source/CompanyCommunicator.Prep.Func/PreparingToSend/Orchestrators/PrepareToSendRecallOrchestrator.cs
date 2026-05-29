// <copyright file="PrepareToSendRecallOrchestrator.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Azure.Functions.Worker;
    using Microsoft.DurableTask;
    using Microsoft.DurableTask.Client;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;

    /// <summary>
    /// Prepare to Send orchestrator.
    ///
    /// This function prepares to send a notification to the target audience.
    ///
    /// Performs following:
    /// 1. Stores the message in sending notification table.
    /// 2. Syncs recipients information to sent notification table.
    /// 3. Creates teams conversation with recipients if required.
    /// 4. Starts Send Queue orchestration.
    /// </summary>
    public static class PrepareToSendRecallOrchestrator
    {
        /// <summary>
        /// This is the durable orchestration method,
        /// which kicks off the preparing to send process.
        /// </summary>
        /// <param name="context">Durable orchestration context.</param>
        /// <param name="log">Logger.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        [Function(FunctionNames.PrepareToSendRecallOrchestrator)]
        public static async Task RunOrchestrator(
            [OrchestrationTrigger] TaskOrchestrationContext context)
        {
            var log = context.CreateReplaySafeLogger(nameof(Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend.FunctionNames));
            var notificationDataEntity = context.GetInput<NotificationDataEntity>();

            if (!context.IsReplaying)
            {
                log.LogInformation($"Start to prepare to send recall the notification {notificationDataEntity.Id}!");
            }

            try
            {
                if (!context.IsReplaying)
                {
                    log.LogInformation("About to store recall message content.");
                }

                await context.CallActivityAsync(FunctionNames.StoreRecallMessageActivity, notificationDataEntity, FunctionSettings.DefaultRetryOptions);

                if (!context.IsReplaying)
                {
                    log.LogInformation("About to send messages to send recall queue.");
                }

                await context.CallSubOrchestratorAsync(FunctionNames.SendRecallQueueOrchestrator, notificationDataEntity, FunctionSettings.DefaultRetryOptions);

                log.LogInformation($"PrepareToSendRecallOrchestrator successfully completed for notification: {notificationDataEntity.Id}!");
            }
            catch (Exception ex)
            {
                var errorMessage = $"PrepareToSendRecallOrchestrator failed for notification: {notificationDataEntity.Id}. Exception Message: {ex.Message}";
                log.LogError(ex, errorMessage);

                await context.CallActivityAsync(FunctionNames.HandleRecallFailureActivity, (notificationDataEntity, ex), FunctionSettings.DefaultRetryOptions);
            }
        }
    }
}

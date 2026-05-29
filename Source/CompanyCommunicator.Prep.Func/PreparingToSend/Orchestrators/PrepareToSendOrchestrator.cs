// <copyright file="PrepareToSendOrchestrator.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System;
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
    public static class PrepareToSendOrchestrator
    {
        /// <summary>
        /// This is the durable orchestration method,
        /// which kicks off the preparing to send process.
        /// </summary>
        /// <param name="context">Durable orchestration context.</param>
        /// <param name="log">Logger.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        [Function(FunctionNames.PrepareToSendOrchestrator)]
        public static async Task RunOrchestrator(
            [OrchestrationTrigger] TaskOrchestrationContext context)
        {
            var log = context.CreateReplaySafeLogger(nameof(Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend.FunctionNames));
            var notificationDataEntity = context.GetInput<NotificationDataEntity>();

            if (!context.IsReplaying)
            {
                log.LogInformation($"Start to prepare to send the notification {notificationDataEntity.Id}!");
            }

            try
            {
                if (!context.IsReplaying)
                {
                    log.LogInformation("About to store message content.");
                }

                await context.CallActivityAsync(FunctionNames.StoreMessageActivity, notificationDataEntity, FunctionSettings.DefaultRetryOptions);

                if (!context.IsReplaying)
                {
                    log.LogInformation("About to sync recipients.");
                }

                await context.CallSubOrchestratorAsync(FunctionNames.SyncRecipientsOrchestrator, notificationDataEntity, FunctionSettings.DefaultRetryOptions);

                if (!context.IsReplaying)
                {
                    log.LogInformation("About to create conversation for recipients if required.");
                }

                await context.CallSubOrchestratorAsync(FunctionNames.TeamsConversationOrchestrator, notificationDataEntity, FunctionSettings.DefaultRetryOptions);

                if (!context.IsReplaying)
                {
                    log.LogInformation("About to send messages to send queue.");
                }

                await context.CallSubOrchestratorAsync(FunctionNames.SendQueueOrchestrator, notificationDataEntity, FunctionSettings.DefaultRetryOptions);

                log.LogInformation($"PrepareToSendOrchestrator successfully completed for notification: {notificationDataEntity.Id}!");
            }
            catch (Exception ex)
            {
                var errorMessage = $"PrepareToSendOrchestrator failed for notification: {notificationDataEntity.Id}. Exception Message: {ex.Message}";
                log.LogError(ex, errorMessage);

                await context.CallActivityAsync(FunctionNames.HandleFailureActivity, (notificationDataEntity, ex), FunctionSettings.DefaultRetryOptions);
            }
        }
    }
}

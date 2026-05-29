// <copyright file="RecallMessageFunction.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Send.Func
{
    using System;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Web;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Azure.Functions.Worker;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Rest;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Helpers;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RecallSentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Resources;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.SendRecallQueue;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Teams;
    using Microsoft.Teams.Apps.CompanyCommunicator.Send.Func.Services;
    using Microsoft.Azure.Storage;
    using Microsoft.Azure.Storage.Blob;
    using Newtonsoft.Json;

    /// <summary>
    /// Azure Function App triggered by messages from a Service Bus queue
    /// Used for sending messages from the bot.
    /// </summary>
    public class RecallMessageFunction
    {
        /// <summary>
        /// This is set to 10 because the default maximum delivery count from the service bus
        /// message queue before the service bus will automatically put the message in the Dead Letter
        /// Queue is 10.
        /// </summary>
        private static readonly int MaxDeliveryCountForDeadLetter = 10;
        private static readonly string AdaptiveCardContentType = "application/vnd.microsoft.card.adaptive";

        private readonly int maxNumberOfAttempts;
        private readonly double sendRetryDelayNumberOfSeconds;
        private readonly INotificationService notificationService;
        private readonly IMessageService messageService;
        private readonly ISendRecallQueue sendRecallQueue;
        private readonly IStringLocalizer<Strings> localizer;
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly CloudStorageHelper cloudStorageHelper;
        private readonly ILogger<RecallMessageFunction> log;

        /// <summary>
        /// Initializes a new instance of the <see cref="RecallMessageFunction"/> class.
        /// </summary>
        public RecallMessageFunction(
            IOptions<SendFunctionOptions> options,
            INotificationService notificationService,
            IMessageService messageService,
            ISendRecallQueue sendRecallQueue,
            IStringLocalizer<Strings> localizer,
            IOptions<RepositoryOptions> optionsRepository,
            ISentNotificationDataRepository sentNotificationDataRepository,
            CloudStorageHelper cloudStorageHelper,
            ILogger<RecallMessageFunction> log)
        {
            if (options is null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (optionsRepository is null)
            {
                throw new ArgumentNullException(nameof(optionsRepository));
            }

            this.maxNumberOfAttempts = options.Value.MaxNumberOfAttempts;
            this.sendRetryDelayNumberOfSeconds = options.Value.SendRetryDelayNumberOfSeconds;

            this.notificationService = notificationService ?? throw new ArgumentNullException(nameof(notificationService));
            this.messageService = messageService ?? throw new ArgumentNullException(nameof(messageService));
            this.sendRecallQueue = sendRecallQueue ?? throw new ArgumentNullException(nameof(sendRecallQueue));
            this.localizer = localizer ?? throw new ArgumentNullException(nameof(localizer));
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentNullException(nameof(sentNotificationDataRepository));
            this.cloudStorageHelper = cloudStorageHelper ?? throw new ArgumentNullException(nameof(cloudStorageHelper));
            this.log = log;
        }

        /// <summary>
        /// Azure Function App triggered by messages from a Service Bus queue
        /// Used for sending messages from the bot.
        /// </summary>
        /// <param name="myQueueItem">The Service Bus queue item.</param>
        /// <param name="deliveryCount">The deliver count.</param>
        /// <param name="enqueuedTimeUtc">The enqueued time.</param>
        /// <param name="messageId">The message ID.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [Function("RecallMessageFunction")]
        public async Task Run(
            [ServiceBusTrigger(
                SendRecallQueue.QueueName,
                Connection = SendRecallQueue.ServiceBusConnectionConfigurationKey)]
            string myQueueItem,
            int deliveryCount,
            DateTime enqueuedTimeUtc,
            string messageId)
        {
            var log = this.log;
            log.LogInformation($"C# ServiceBus recall queue trigger function processed message: {myQueueItem}");

            var messageContent = JsonConvert.DeserializeObject<SendRecallQueueMessageContent>(myQueueItem);

            try
            {
                // Check if notification is pending.
                var isPending = await this.notificationService.IsPendingRecallNotification(messageContent);
                if (!isPending)
                {
                    // Notification is either already sent or failed and shouldn't be retried.
                    return;
                }

                // Check if conversationId is set to send message.
                if (string.IsNullOrWhiteSpace(messageContent.RecipientData.ConversationId))
                {
                    await this.notificationService.UpdateSentRecallNotification(
                        notificationId: messageContent.NotificationId,
                        recipientId: messageContent.RecipientData.RecipientId,
                        totalNumberOfSendThrottles: 0,
                        statusCode: RecallSentNotificationDataEntity.FinalFaultedStatusCode,
                        allSendStatusCodes: $"{RecallSentNotificationDataEntity.FinalFaultedStatusCode},",
                        errorMessage: this.localizer.GetString("AppNotInstalled"),
                        activityId: string.Empty);
                    return;
                }

                // Check if the system is throttled.
                var isThrottled = await this.notificationService.IsSendRecallNotificationThrottled();
                if (isThrottled)
                {
                    // Re-Queue with delay.
                    await this.sendRecallQueue.SendDelayedAsync(messageContent, this.sendRetryDelayNumberOfSeconds);
                    return;
                }

                // Send message.
                var messageActivity = await this.GetMessageActivity(messageContent, messageContent.RecipientData.RecipientId, log);
                messageActivity.ReplyToId = messageContent.RecipientData.ActivityId;
                messageActivity.Id = messageContent.RecipientData.ActivityId;

                var response = await this.messageService.UpdateMessageAsync(
                    message: messageActivity,
                    serviceUrl: messageContent.RecipientData.ServiceUrl,
                    conversationId: messageContent.RecipientData.ConversationId,
                    maxAttempts: this.maxNumberOfAttempts,
                    logger: log);

                // Process response.
                await this.ProcessResponseAsync(messageContent, response, log);
            }
            catch (InvalidOperationException exception)
            {
                // Bad message shouldn't be requeued.
                log.LogError(exception, $"InvalidOperationException thrown. Error message: {exception.Message}");
            }
            catch (Exception e)
            {
                var errorMessage = $"{e.GetType()}: {e.Message}";
                log.LogError(e, $"Failed to send message. ErrorMessage: {errorMessage}");

                // Update status code depending on delivery count.
                var statusCode = RecallSentNotificationDataEntity.FaultedAndRetryingStatusCode;
                if (deliveryCount >= RecallMessageFunction.MaxDeliveryCountForDeadLetter)
                {
                    // Max deliveries attempted. No further retries.
                    statusCode = RecallSentNotificationDataEntity.FinalFaultedStatusCode;
                }

                // Update sent notification table.
                await this.notificationService.UpdateSentRecallNotification(
                    notificationId: messageContent.NotificationId,
                    recipientId: messageContent.RecipientData.RecipientId,
                    totalNumberOfSendThrottles: 0,
                    statusCode: statusCode,
                    allSendStatusCodes: $"{statusCode},",
                    errorMessage: errorMessage,
                    activityId: string.Empty);

                throw;
            }
        }

        /// <summary>
        /// Process send notification response.
        /// </summary>
        /// <param name="messageContent">Message content.</param>
        /// <param name="sendMessageResponse">Send notification response.</param>
        /// <param name="log">Logger.</param>
        private async Task ProcessResponseAsync(
            SendRecallQueueMessageContent messageContent,
            SendMessageResponse sendMessageResponse,
            ILogger log)
        {
            if (sendMessageResponse.ResultType == SendMessageResult.Succeeded)
            {
                log.LogInformation($"Successfully recall the message." +
                    $"\nRecipient Id: {messageContent.RecipientData.RecipientId}");

                // Delete message from send notification data.
                try
                {
                    var result = await this.sentNotificationDataRepository.GetAsync(messageContent.NotificationId, messageContent.RecipientData.RecipientId);

                    if (result != null)
                    {
                        await this.sentNotificationDataRepository.DeleteAsync(result);
                    }
                }
                catch (Exception ex)
                {
                    log.LogError($"Failed to delete message from sent notification." +
                    $"\nReceipient Id: {messageContent.NotificationId}" +
                    $"\nNotification Id: {messageContent.RecipientData.RecipientId}" +
                    $"\nResult: {ex.InnerException.Message}." +
                    $"\nErrorMessage: {ex.Message}.");
                }

                /*
                // Delete message from Reaction data if any
                try
                {
                    await this.cloudStorageHelper.DeleteReactionDataAsync(messageContent.NotificationId, messageContent.RecipientData.RowKey);
                }
                catch (Exception ex)
                {
                    log.LogError($"Failed to delete message from reaction data." +
                    $"\nReceipient Id: {messageContent.NotificationId}" +
                    $"\nNotification Id: {messageContent.RecipientData.RecipientId}" +
                    $"\nResult: {ex.InnerException.Message}." +
                    $"\nErrorMessage: {ex.Message}.");
                }

                // Delete message from survey answer data if any
                try
                {
                    await this.cloudStorageHelper.DeleteQuestionAnswerDataAsync(messageContent.NotificationId, messageContent.RecipientData.RowKey);
                }
                catch (Exception ex)
                {
                    log.LogError($"Failed to delete message from question and answer data." +
                    $"\nReceipient Id: {messageContent.NotificationId}" +
                    $"\nNotification Id: {messageContent.RecipientData.RecipientId}" +
                    $"\nResult: {ex.InnerException.Message}." +
                    $"\nErrorMessage: {ex.Message}.");
                }*/
            }
            else
            {
                log.LogError($"Failed to recall message." +
                    $"\nRecipient Id: {messageContent.RecipientData.RecipientId}" +
                    $"\nResult: {sendMessageResponse.ResultType}." +
                    $"\nErrorMessage: {sendMessageResponse.ErrorMessage}.");
            }

            await this.notificationService.UpdateSentRecallNotification(
                    notificationId: messageContent.NotificationId,
                    recipientId: messageContent.RecipientData.RecipientId,
                    totalNumberOfSendThrottles: sendMessageResponse.TotalNumberOfSendThrottles,
                    statusCode: sendMessageResponse.StatusCode,
                    allSendStatusCodes: sendMessageResponse.AllSendStatusCodes,
                    errorMessage: sendMessageResponse.ErrorMessage,
                    activityId: sendMessageResponse.ActivityId);

            // Throttled
            if (sendMessageResponse.ResultType == SendMessageResult.Throttled)
            {
                // Set send function throttled.
                await this.notificationService.SetSendNotificationThrottled(this.sendRetryDelayNumberOfSeconds);

                // Requeue.
                await this.sendRecallQueue.SendDelayedAsync(messageContent, this.sendRetryDelayNumberOfSeconds);
                return;
            }
        }

        private async Task<IMessageActivity> GetMessageActivity(SendRecallQueueMessageContent message, string recipientId, ILogger log)
        {
            var content = @"{
    ""type"": ""AdaptiveCard"",
    ""body"": [
        {
                ""type"": ""TextBlock"",
            ""weight"": ""Bolder"",
            ""text"": """ + message.MessageTitle + @""",
            ""size"": ""Large"",
            ""wrap"": true
        },
        {
                ""type"": ""TextBlock"",
            ""weight"": ""Default"",
            ""text"": ""This message has been recalled by author."",
            ""size"": ""Default"",
            ""wrap"": true
        }
    ],
    ""msteams"": {
                ""width"": ""Full""
    },
    ""$schema"": ""http://adaptivecards.io/schemas/adaptive-card.json"",
    ""version"": ""1.4""
}";
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCardContentType,
                Content = JsonConvert.DeserializeObject(content),
            };
            await Task.Delay(0);
            return MessageFactory.Attachment(adaptiveCardAttachment);
        }
    }
}

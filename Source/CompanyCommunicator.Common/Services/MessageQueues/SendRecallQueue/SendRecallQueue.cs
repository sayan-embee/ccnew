// <copyright file="SendRecallQueue.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.SendRecallQueue
{
    using Microsoft.Extensions.Options;

    /// <summary>
    /// The message queue service connected to the "company-communicator-send-recall" queue in Azure service bus.
    /// </summary>
    public class SendRecallQueue : BaseQueue<SendRecallQueueMessageContent>, ISendRecallQueue
    {
        /// <summary>
        /// Queue name of the send queue.
        /// </summary>
        public const string QueueName = "company-communicator-send-recall";

        /// <summary>
        /// Initializes a new instance of the <see cref="SendRecallQueue"/> class.
        /// </summary>
        /// <param name="messageQueueOptions">The message queue options.</param>
        public SendRecallQueue(IOptions<MessageQueueOptions> messageQueueOptions)
            : base(
                  serviceBusConnectionString: messageQueueOptions.Value.ServiceBusConnection,
                  queueName: SendRecallQueue.QueueName)
        {
        }
    }
}

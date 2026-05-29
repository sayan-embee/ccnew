// <copyright file="PrepareToRecallQueue.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.PrepareToRecallQueue
{
    using Microsoft.Extensions.Options;

    /// <summary>
    /// The message queue service connected to the "company-communicator-prep-recall" queue in Azure service bus.
    /// </summary>
    public class PrepareToRecallQueue : BaseQueue<PrepareToRecallQueueMessageContent>, IPrepareToRecallQueue
    {
        /// <summary>
        /// Queue name of the prepare to send queue.
        /// </summary>
        public const string QueueName = "company-communicator-prep-recall";

        /// <summary>
        /// Initializes a new instance of the <see cref="PrepareToRecallQueue"/> class.
        /// </summary>
        /// <param name="messageQueueOptions">The message queue options.</param>
        public PrepareToRecallQueue(IOptions<MessageQueueOptions> messageQueueOptions)
            : base(
                  serviceBusConnectionString: messageQueueOptions.Value.ServiceBusConnection,
                  queueName: PrepareToRecallQueue.QueueName)
        {
        }
    }
}

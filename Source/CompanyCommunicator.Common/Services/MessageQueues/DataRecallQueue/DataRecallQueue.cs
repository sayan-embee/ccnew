// <copyright file="DataRecallQueue.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.DataRecallQueue
{
    using Microsoft.Extensions.Options;

    /// <summary>
    /// The message queue service connected to the "company-communicator-data-recall" queue in Azure service bus.
    /// </summary>
    public class DataRecallQueue : BaseQueue<DataRecallQueueMessageContent>, IDataRecallQueue
    {
        /// <summary>
        /// Queue name of the data queue.
        /// </summary>
        public const string QueueName = "company-communicator-data-recall";

        /// <summary>
        /// Initializes a new instance of the <see cref="DataRecallQueue"/> class.
        /// </summary>
        /// <param name="messageQueueOptions">The message queue options.</param>
        public DataRecallQueue(IOptions<MessageQueueOptions> messageQueueOptions)
            : base(
                  serviceBusConnectionString: messageQueueOptions.Value.ServiceBusConnection,
                  queueName: DataRecallQueue.QueueName)
        {
        }
    }
}

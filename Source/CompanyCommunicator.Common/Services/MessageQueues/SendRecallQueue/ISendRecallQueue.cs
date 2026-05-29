// <copyright file="ISendRecallQueue.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.SendRecallQueue
{
    /// <summary>
    /// interface for Send Queue.
    /// </summary>
    public interface ISendRecallQueue : IBaseQueue<SendRecallQueueMessageContent>
    {
    }
}

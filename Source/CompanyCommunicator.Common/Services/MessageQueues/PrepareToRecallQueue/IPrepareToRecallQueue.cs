// <copyright file="IPrepareToRecallQueue.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.PrepareToRecallQueue
{
    /// <summary>
    /// interface for Prepare to send Queue.
    /// </summary>
    public interface IPrepareToRecallQueue : IBaseQueue<PrepareToRecallQueueMessageContent>
    {
    }
}

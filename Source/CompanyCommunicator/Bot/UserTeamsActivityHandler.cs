// <copyright file="UserTeamsActivityHandler.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Bot
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Helpers;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Models;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Company Communicator User Bot.
    /// Captures user data, team data.
    /// </summary>
    public class UserTeamsActivityHandler : TeamsActivityHandler
    {
        private static readonly string TeamRenamedEventType = "teamRenamed";

        private readonly TeamsDataCapture teamsDataCapture;

        private readonly CloudStorageHelper cloudStorageHelper;

        private readonly ISentNotificationDataRepository sentNotificationDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserTeamsActivityHandler"/> class.
        /// </summary>
        /// <param name="teamsDataCapture">Teams data capture service.</param>
        /// <param name="cloudStorageHelper">Cloud storage helper service.</param>
        /// <param name="sentNotificationDataRepository">sent notificaton data repository service.</param>
        public UserTeamsActivityHandler(TeamsDataCapture teamsDataCapture, CloudStorageHelper cloudStorageHelper, ISentNotificationDataRepository sentNotificationDataRepository)
        {
            this.teamsDataCapture = teamsDataCapture ?? throw new ArgumentNullException(nameof(teamsDataCapture));
            this.cloudStorageHelper = cloudStorageHelper ?? throw new ArgumentNullException(nameof(cloudStorageHelper));
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentNullException(nameof(sentNotificationDataRepository));
        }

        /// <summary>
        /// Invoked when a conversation update activity is received from the channel.
        /// </summary>
        /// <param name="turnContext">The context object for this turn.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects
        /// or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnConversationUpdateActivityAsync(
            ITurnContext<IConversationUpdateActivity> turnContext,
            CancellationToken cancellationToken)
        {
            // base.OnConversationUpdateActivityAsync is useful when it comes to responding to users being added to or removed from the conversation.
            // For example, a bot could respond to a user being added by greeting the user.
            // By default, base.OnConversationUpdateActivityAsync will call <see cref="OnMembersAddedAsync(IList{ChannelAccount}, ITurnContext{IConversationUpdateActivity}, CancellationToken)"/>
            // if any users have been added or <see cref="OnMembersRemovedAsync(IList{ChannelAccount}, ITurnContext{IConversationUpdateActivity}, CancellationToken)"/>
            // if any users have been removed. base.OnConversationUpdateActivityAsync checks the member ID so that it only responds to updates regarding members other than the bot itself.
            await base.OnConversationUpdateActivityAsync(turnContext, cancellationToken);

            var activity = turnContext.Activity;

            var isTeamRenamed = this.IsTeamInformationUpdated(activity);
            if (isTeamRenamed)
            {
                await this.teamsDataCapture.OnTeamInformationUpdatedAsync(activity);
            }

            if (activity.MembersAdded != null)
            {
                var teamsChannelAccount = await TeamsInfo.GetMemberAsync(turnContext, turnContext.Activity.From.Id, cancellationToken);
                await this.teamsDataCapture.OnBotAddedAsync(activity, teamsChannelAccount);
            }

            if (activity.MembersRemoved != null)
            {
                await this.teamsDataCapture.OnBotRemovedAsync(activity);
            }
        }

        /// <inheritdoc/>
        protected override async Task OnReactionsAddedAsync(IList<MessageReaction> messageReactions, ITurnContext<IMessageReactionActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                foreach (var reaction in messageReactions)
                {
                    IMessageReactionActivity activity = turnContext.Activity;
                    var member = activity.From;
                    var member1 = await TeamsInfo.GetMemberAsync(turnContext, turnContext.Activity.From.Id, cancellationToken);
                    UserReaction userreaction = new UserReaction();
                    userreaction.AadId = member.AadObjectId;
                    userreaction.FromId = member.Id;
                    userreaction.Name = member1.Name;
                    userreaction.Email = member1.Email;
                    userreaction.TenantId = member1.TenantId;
                    userreaction.ReactionType = reaction.Type;
                    userreaction.ActivityId = turnContext.Activity.Id;
                    userreaction.ReactionDate = DateTime.UtcNow;

                    string strFilter = TableQuery.GenerateFilterCondition("ActivityId", QueryComparisons.Equal, userreaction.ActivityId);

                    var notificationEntities = await this.sentNotificationDataRepository.GetWithFilterCustomAsync(strFilter);
                    if (notificationEntities.Any())
                    {
                        userreaction.NotificationId = notificationEntities.First().PartitionKey;
                    }

                    string valdata = await this.cloudStorageHelper.MergeUserReactionData(userreaction);
                }
            }
            catch (Exception ex)
            {
                var exMessage = ex.Message;
            }
        }

        /// <inheritdoc/>
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));

                // JToken commandToken = null;
                IMessageActivity activity = turnContext.Activity;
                var member = activity.From;
                string aadid = member.AadObjectId;
                string name = member.Name;
                string id = member.Id;
                if (turnContext.Activity.Value != null)
                {
                    var member1 = await TeamsInfo.GetMemberAsync(turnContext, turnContext.Activity.From.Id, cancellationToken);

                    // commandToken = JToken.Parse(turnContext.Activity.Value.ToString());
                    QuestionAnswer questiondata = JsonConvert.DeserializeObject<QuestionAnswer>(turnContext.Activity.Value.ToString());
                    questiondata.FromId = id;
                    if (!string.IsNullOrEmpty(name))
                    {
                        name = name.Replace("/", "-");
                        name = name.Replace("\\", "-");
                        name = name.Replace("#", "-");
                        name = name.Replace("?", "-");
                    }

                    questiondata.Name = name;
                    questiondata.AadId = aadid;
                    questiondata.Email = member1.Email;
                    questiondata.TenantId = member1.TenantId;
                    questiondata.SubmittedOn = DateTime.UtcNow;

                    string valdata = await this.cloudStorageHelper.MergeAdaptiveCardData(questiondata);
                    var adaptiveCardAttachments = new Attachment()
                    {
                        ContentType = "application/vnd.microsoft.card.adaptive",
                        Content = JsonConvert.DeserializeObject(valdata),
                    };

                    var appmessage = Activity.CreateMessageActivity();
                    appmessage.Attachments.Add(adaptiveCardAttachments);
                    appmessage.Id = turnContext.Activity.ReplyToId;
                    var approverCard = await turnContext.UpdateActivityAsync(appmessage, cancellationToken);
                }
            }
            catch (Exception ex)
            {
                var exMessage = ex.Message;
                await turnContext.SendActivityAsync(MessageFactory.Text(exMessage, exMessage), cancellationToken);
            }
        }

        /// <summary>
        /// Invoked when a conversation update activity is received from the channel.
        /// </summary>
        /// <param name="membersAdded">Then content object.</param>
        /// <param name="turnContext">The context object for this turn.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects
        /// or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var activity = turnContext.Activity;
            if (activity.MembersAdded != null)
            {
                await this.teamsDataCapture.OnBotAddedAsync(activity);
            }
        }

        private bool IsTeamInformationUpdated(IConversationUpdateActivity activity)
        {
            if (activity == null)
            {
                return false;
            }

            var channelData = activity.GetChannelData<TeamsChannelData>();
            if (channelData == null)
            {
                return false;
            }

            return UserTeamsActivityHandler.TeamRenamedEventType.Equals(channelData.EventType, StringComparison.OrdinalIgnoreCase);
        }
    }
}
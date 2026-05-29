// <copyright file="CloudStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Cloud storage helper.
    /// </summary>
    public class CloudStorageHelper1
    {
        /// <summary>
        /// Repository option reference.
        /// </summary>
        private readonly IOptions<RepositoryOptions> repositoryOptions;

        /// <summary>
        /// ILogger object reference.
        /// </summary>
        private readonly ILogger<CloudStorageHelper1> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="CloudStorageHelper1"/> class.
        /// </summary>
        /// <param name="logger">Logger.</param>
        /// <param name="repositoryOptions">Repository.</param>
        public CloudStorageHelper1(
            ILogger<CloudStorageHelper1> logger,
            IOptions<RepositoryOptions> repositoryOptions)
        {
            this.repositoryOptions = repositoryOptions ?? throw new ArgumentNullException(nameof(repositoryOptions));
            this.logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// Merge Adaptive card data.
        /// </summary>
        /// <param name="formData">Question answer form data.</param>
        /// <returns>json string.</returns>
        public async Task<string> MergeAdaptiveCardData(QuestionAnswer formData)
        {
            try
            {
                int maxQuestion = 8;
                var storageConnectionString = this.repositoryOptions.Value.StorageAccountConnectionString;
                string body = "{" +
                          "\"type\": \"TextBlock\"," +
                          "\"text\": \"" + formData.Title + "\"," +
                          "\"size\": \"ExtraLarge\"," +
                          "\"wrap\": true," +
                          "\"weight\": \"Bolder\"" +
                        "}," +
                        "{" +
                          "\"type\": \"TextBlock\"," +
                          "\"text\": \"" + formData.Author + "\"," +
                          "\"size\": \"Small\"," +
                          "\"wrap\": true," +
                          "\"weight\": \"Lighter\"" +
                        "},";

                // string question0 = string.Empty, question1 = string.Empty, question2 = string.Empty, question3 = string.Empty, question4 = string.Empty, question5 = string.Empty;
                if (formData.Questions != string.Empty)
                {
                    string[] questionlist = formData.Questions.Split("||");
                    if (questionlist.Length > 0)
                    {
                        for (var i = 0; i < questionlist.Length; i++)
                        {
                            if (maxQuestion >= 8)
                            {
                                break;
                            }

                            var ans = string.Empty;
                            switch (i)
                            {
                                case 0: ans = formData.Answer0; break;
                                case 1: ans = formData.Answer1; break;
                                case 2: ans = formData.Answer2; break;
                                case 3: ans = formData.Answer3; break;
                                case 4: ans = formData.Answer4; break;
                                case 5: ans = formData.Answer5; break;
                                case 6: ans = formData.Answer6; break;
                                case 7: ans = formData.Answer7; break;
                                default: break;
                            }

                            body += "{" +
                          "\"type\": \"TextBlock\"," +
                          "\"text\": \"" + (i + 1) + "." + questionlist[i] + "\"," +
                          "\"size\": \"Medium\"," +
                          "\"wrap\": true," +
                          "\"horizontalAlignment\": \"Left\"" +
                        "}," +
                        "{" +
                          "\"type\": \"TextBlock\"," +
                          "\"text\": \"Ans: " + ans + "\"," +
                          "\"size\": \"Medium\"," +
                          "\"wrap\": true," +
                          "\"horizontalAlignment\": \"Left\"" +
                        "}";
                        }
                    }

                    // if (questionlist.Length > 0)
                    // {
                    //    question0 = questionlist[0];
                    //    body += "{" +
                    //      "\"type\": \"TextBlock\"," +
                    //      "\"text\": \"1." + question0 + "\"," +
                    //      "\"size\": \"Medium\"," +
                    //      "\"wrap\": true," +
                    //      "\"horizontalAlignment\": \"Left\"" +
                    //    "}," +
                    //    "{" +
                    //      "\"type\": \"TextBlock\"," +
                    //      "\"text\": \"Ans: " + formData.answer0 + "\"," +
                    //      "\"size\": \"Medium\"," +
                    //      "\"wrap\": true," +
                    //      "\"horizontalAlignment\": \"Left\"" +
                    //    "}";
                    // }

                    // if (questionlist.Length > 1)
                    // {
                    //    question1 = questionlist[1];
                    //    body += ",{" +
                    //      "\"type\": \"TextBlock\"," +
                    //      "\"text\": \"2." + question1 + "\"," +
                    //      "\"size\": \"Medium\"," +
                    //      "\"wrap\": true," +
                    //      "\"horizontalAlignment\": \"Left\"" +
                    //    "}," +
                    //    "{" +
                    //      "\"type\": \"TextBlock\"," +
                    //      "\"text\": \"Ans: " + formData.answer1 + "\"," +
                    //      "\"size\": \"Medium\"," +
                    //      "\"wrap\": true," +
                    //      "\"horizontalAlignment\": \"Left\"" +
                    //    "}";
                    // }

                    // if (questionlist.Length > 2)
                    // {
                    //    question2 = questionlist[2];
                    //    body += ",{" +
                    //      "\"type\": \"TextBlock\"," +
                    //      "\"text\": \"3." + question2 + "\"," +
                    //      "\"size\": \"Medium\"," +
                    //      "\"wrap\": true," +
                    //      "\"horizontalAlignment\": \"Left\"" +
                    //    "}," +
                    //    "{" +
                    //      "\"type\": \"TextBlock\"," +
                    //      "\"text\": \"Ans: " + formData.answer2 + "\"," +
                    //      "\"size\": \"Medium\"," +
                    //      "\"wrap\": true," +
                    //      "\"horizontalAlignment\": \"Left\"" +
                    //    "}";
                    // }

                    // if (questionlist.Length > 3)
                    // {
                    //    question3 = questionlist[3];
                    //    body += ",{" +
                    //      "\"type\": \"TextBlock\"," +
                    //      "\"text\": \"4." + question3 + "\"," +
                    //      "\"size\": \"Medium\"," +
                    //      "\"wrap\": true," +
                    //      "\"horizontalAlignment\": \"Left\"" +
                    //    "}," +
                    //    "{" +
                    //      "\"type\": \"TextBlock\"," +
                    //      "\"text\": \"Ans: " + formData.answer3 + "\"," +
                    //      "\"size\": \"Medium\"," +
                    //      "\"wrap\": true," +
                    //      "\"horizontalAlignment\": \"Left\"" +
                    //    "}";
                    // }

                    // if (questionlist.Length > 4)
                    // {
                    //    question4 = questionlist[4];
                    //    body += ",{" +
                    //      "\"type\": \"TextBlock\"," +
                    //      "\"text\": \"5." + question4 + "\"," +
                    //      "\"size\": \"Medium\"," +
                    //      "\"wrap\": true," +
                    //      "\"horizontalAlignment\": \"Left\"" +
                    //    "}," +
                    //    "{" +
                    //      "\"type\": \"TextBlock\"," +
                    //      "\"text\": \"Ans: " + formData.answer5 + "\"," +
                    //      "\"size\": \"Medium\"," +
                    //      "\"wrap\": true," +
                    //      "\"horizontalAlignment\": \"Left\"" +
                    //    "}";
                    // }

                    // if (questionlist.Length > 5)
                    // {
                    //    body += ",{" +
                    //      "\"type\": \"TextBlock\"," +
                    //      "\"text\": \"6." + question5 + "\"," +
                    //      "\"size\": \"Medium\"," +
                    //      "\"wrap\": true," +
                    //      "\"horizontalAlignment\": \"Left\"" +
                    //    "}," +
                    //    "{" +
                    //      "\"type\": \"TextBlock\"," +
                    //      "\"text\": \"Ans: " + formData.answer5 + "\"," +
                    //      "\"size\": \"Medium\"," +
                    //      "\"wrap\": true," +
                    //      "\"horizontalAlignment\": \"Left\"" +
                    //    "}";
                    // }
                    QuestionAnswerAdaptiveCardEntity adaptiveCard = new QuestionAnswerAdaptiveCardEntity(formData.NotificationId, formData.Name)
                    {
                        NotificationId = formData.NotificationId,
                        Title = formData.Title,
                        Author = formData.Author,
                        Email = formData.Email,
                        TenantId = formData.TenantId,
                        SubmittedOn = formData.SubmittedOn,
                        ConversationId = formData.ConversationId,
                        ActivityId = formData.ActivityId,
                    };
                    if (questionlist.Length > 0)
                    {
                        for (var i = 0; i < questionlist.Length; i++)
                        {
                            switch (i)
                            {
                                case 0: adaptiveCard.Question0 = questionlist[i]; adaptiveCard.Answer0 = formData.Answer0; break;
                                case 1: adaptiveCard.Question1 = questionlist[i]; adaptiveCard.Answer1 = formData.Answer1; break;
                                case 2: adaptiveCard.Question2 = questionlist[i]; adaptiveCard.Answer2 = formData.Answer2; break;
                                case 3: adaptiveCard.Question3 = questionlist[i]; adaptiveCard.Answer3 = formData.Answer3; break;
                                case 4: adaptiveCard.Question4 = questionlist[i]; adaptiveCard.Answer4 = formData.Answer4; break;
                                case 5: adaptiveCard.Question5 = questionlist[i]; adaptiveCard.Answer5 = formData.Answer5; break;
                                case 6: adaptiveCard.Question6 = questionlist[i]; adaptiveCard.Answer6 = formData.Answer6; break;
                                case 7: adaptiveCard.Question7 = questionlist[i]; adaptiveCard.Answer7 = formData.Answer7; break;
                                default: break;
                            }
                        }
                    }

                    var tableName = "QuestionAnswer";
                    CloudStorageAccount storageAccount;
                    storageAccount = CloudStorageAccount.Parse(storageConnectionString);

                    CloudTableClient tableClient = storageAccount.CreateCloudTableClient(new TableClientConfiguration());
                    CloudTable table = tableClient.GetTableReference(tableName);
                    await table.CreateIfNotExistsAsync();
                    TableOperation insertOnMergeOperation = TableOperation.InsertOrMerge(adaptiveCard);
                    TableResult result = await table.ExecuteAsync(insertOnMergeOperation);
                }

                return this.ReturnAdaptiveCardJSON(body);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error in gettting MergeAdaptiveCardData for notification id : {formData.NotificationId}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Return Adaptive card Json object.
        /// </summary>
        /// <param name="body">String.</param>
        /// <returns>Json String.</returns>
        public string ReturnAdaptiveCardJSON(string body)
        {
            string jsonString = "{" +
              "\"type\": \"AdaptiveCard\"," +
              "\"body\": [" + body +
              "]," +
              "\"$schema\": \"http://adaptivecards.io/schemas/adaptive-card.json\"," +
              "\"version\": \"1.2\"" +
            "}";
            return jsonString;
        }

        /// <summary>
        /// Insert or Merge User Reaction data.
        /// </summary>
        /// <param name="formData">User Reaction data.</param>
        /// <returns>String.</returns>
        public async Task<string> MergeUserReactionData(UserReaction formData)
        {
            try
            {
                var storageConnectionString = this.repositoryOptions.Value.StorageAccountConnectionString;
                var tableName = "UserReaction";
                CloudStorageAccount storageAccount;
                storageAccount = CloudStorageAccount.Parse(storageConnectionString);
                UserReactionEntity entity = new UserReactionEntity(formData.ActivityId, formData.AadId)
                {
                    ActivityId = formData.ActivityId,
                    FromId = formData.FromId,
                    Name = formData.Name,
                    AadId = formData.AadId,
                    ReactionType = formData.ReactionType,
                    Email = formData.Email,
                    TenantId = formData.TenantId,
                    ReactionDate = formData.ReactionDate,
                    NotificationId = formData.NotificationId,
                };

                CloudTableClient tableClient = storageAccount.CreateCloudTableClient(new TableClientConfiguration());
                CloudTable table = tableClient.GetTableReference(tableName);
                await table.CreateIfNotExistsAsync();
                TableOperation insertOnMergeOperation = TableOperation.InsertOrMerge(entity);
                TableResult result = await table.ExecuteAsync(insertOnMergeOperation);
                return "OK";
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error in gettting MergeUserReactionData for notification id : {formData.AadId}");
                return string.Empty;
            }
}

        /// <summary>
        /// Get list of Survey quenstion and answer.
        /// </summary>
        /// <param name="notificationId">Notificatoin Id.</param>
        /// <returns>List of Question Answer.</returns>
        public async Task<IEnumerable<QuestionAnswerExport>> GetSurveryList(string notificationId)
        {
            IEnumerable<QuestionAnswerExport> dataentity = new List<QuestionAnswerExport>();
            try
            {
                await Task.Delay(0);

                var storageConnectionString = this.repositoryOptions.Value.StorageAccountConnectionString;
                var tableName = "QuestionAnswer";
                CloudStorageAccount storageAccount;
                storageAccount = CloudStorageAccount.Parse(storageConnectionString);

                CloudTableClient tableClient = storageAccount.CreateCloudTableClient(new TableClientConfiguration());
                CloudTable table = tableClient.GetTableReference(tableName);
                var entities = table.ExecuteQuery(new TableQuery<QuestionAnswerAdaptiveCardEntity>()).ToList();
                dataentity = from s in entities
                                 where s.PartitionKey == notificationId
                                 select new QuestionAnswerExport
                                 {
                                     QuestionTitle = s.Title,
                                     Author = s.Author,
                                     Name = s.RowKey,
                                     Question1 = s.Question0,
                                     Answer1 = Convert.ToString(s.Answer0),
                                     Question2 = s.Question1,
                                     Answer2 = Convert.ToString(s.Answer1),
                                     Question3 = s.Question2,
                                     Answer3 = Convert.ToString(s.Answer2),
                                     Question4 = s.Question3,
                                     Answer4 = Convert.ToString(s.Answer3),
                                     Question5 = s.Question4,
                                     Answer5 = Convert.ToString(s.Answer4),
                                     Question6 = s.Question5,
                                     Answer6 = Convert.ToString(s.Answer5),
                                 };
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error in gettting GetSurveryList for notification id : {notificationId}");
            }

            return dataentity;
        }

        /// <summary>
        /// Get reaction list.
        /// </summary>
        /// <param name="notificationId">Notification Id.</param>
        /// <returns>Reaction list data.</returns>
        public async Task<UserReactionExport> GetReactionList(string notificationId)
        {
            UserReactionExport exportData = new UserReactionExport();
            try
            {
                int reactionLike = 0;
                int reactionHeart = 0;
                int reactionLaugh = 0;
                int reactionSurprised = 0;
                int reactionSad = 0;
                int reactionAngry = 0;
                var storageConnectionString = this.repositoryOptions.Value.StorageAccountConnectionString;
                var tableName = "SentNotificationData";
                CloudStorageAccount storageAccount;
                storageAccount = CloudStorageAccount.Parse(storageConnectionString);

                CloudTableClient tableClient = storageAccount.CreateCloudTableClient(new TableClientConfiguration());
                CloudTable table = tableClient.GetTableReference(tableName);
                var entities = table.ExecuteQuery(new TableQuery<SentNotificationDataEntity>()).ToList().Where(x => x.PartitionKey == notificationId);
                foreach (var entity in entities)
                {
                    // int like = 0, heart = 0,laugh=0,surprised=0,
                    // int sad = 0, angry = 0;
                    CloudTable table1 = tableClient.GetTableReference("UserReaction");
                    var subentities = table1.ExecuteQuery(new TableQuery<UserReactionEntity>()).ToList().Where(x => x.ActivityId == entity.ActivityId);
                    foreach (var subentity in subentities)
                    {
                        if (subentity.ReactionType.ToLower() == "like")
                        {
                            reactionLike++;
                        }
                        else if (subentity.ReactionType.ToLower() == "heart")
                        {
                            reactionHeart++;
                        }
                        else if (subentity.ReactionType.ToLower() == "laugh")
                        {
                            reactionLaugh++;
                        }
                        else if (subentity.ReactionType.ToLower() == "surprised")
                        {
                            reactionSurprised++;
                        }
                        else if (subentity.ReactionType.ToLower() == "sad")
                        {
                            reactionSad++;
                        }
                        else if (subentity.ReactionType.ToLower() == "angry")
                        {
                            reactionAngry++;
                        }
                    }

                    // ReactionLike = ReactionLike + like;
                    // ReactionHeart = ReactionHeart + heart;
                    // ReactionLaugh = ReactionHeart + laugh;
                    // ReactionSurprised = ReactionHeart + surprised;
                    // ReactionSad = ReactionHeart + sad;
                    // ReactionAngry = ReactionHeart + angry;
                }

                exportData.LikeCount = reactionLike;
                exportData.HeartCount = reactionHeart;
                exportData.LaughCount = reactionLaugh;
                exportData.SurprisedCount = reactionSurprised;
                exportData.SadCount = reactionSad;
                exportData.AngryCount = reactionAngry;

                await Task.Delay(0);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error in gettting GetReactionList for notification id : {notificationId}");
            }

            return exportData;
        }

        /// <summary>
        /// Get the list of user along with their read receipt and reaction.
        /// </summary>
        /// <param name="notificationId">Notification Id.</param>
        /// <returns>Sent Notification List.</returns>
        public async Task<List<ReportSentNotificationDetailModel>> GetReactionListDetail(string notificationId)
        {
            List<ReportSentNotificationDetailModel> exportData = new List<ReportSentNotificationDetailModel>();
            try
            {
                var storageConnectionString = this.repositoryOptions.Value.StorageAccountConnectionString;
                CloudStorageAccount storageAccount;
                storageAccount = CloudStorageAccount.Parse(storageConnectionString);

                CloudTableClient tableClient = storageAccount.CreateCloudTableClient(new TableClientConfiguration());
                CloudTable tableSentNotification = tableClient.GetTableReference("SentNotificationData");
                CloudTable tableNotificationData = tableClient.GetTableReference("NotificationData");
                CloudTable tableUserReaction = tableClient.GetTableReference("UserReaction");

                List<Task> lstTask = new List<Task>();
                IEnumerable<SentNotificationDataEntity> entitiesSentNotification = null;
                IEnumerable<NotificationDataEntity> entitiesNotificationData = null;
                IEnumerable<UserReactionEntity> entitiesReactions = null;
                lstTask.Add(Task.Run(() => { entitiesSentNotification = tableSentNotification.ExecuteQuery(new TableQuery<SentNotificationDataEntity>()).ToList().Where(x => x.PartitionKey == notificationId); }));
                lstTask.Add(Task.Run(() => { entitiesNotificationData = tableNotificationData.ExecuteQuery(new TableQuery<NotificationDataEntity>()).ToList().Where(x => x.PartitionKey == "SentNotifications" && x.RowKey == notificationId); }));
                lstTask.Add(Task.Run(() => { entitiesReactions = tableUserReaction.ExecuteQuery(new TableQuery<UserReactionEntity>()).ToList().Where(x => x.PartitionKey == notificationId); }));

                Task.WaitAll(lstTask.ToArray());

                if (entitiesNotificationData != null && entitiesSentNotification != null)
                {
                    foreach (var entity in entitiesSentNotification)
                    {
                        UserReactionEntity reaction = null;

                        if (entitiesReactions != null && entitiesReactions.Any())
                        {
                            reaction = entitiesReactions.Where(x => x.ActivityId == entity.ActivityId).First();
                        }

                        var reportModel = new ReportSentNotificationDetailModel();
                        reportModel.Name = entity.UserId;
                        reportModel.Email = entity.Email;
                        reportModel.NotificationId = entity.PartitionKey;
                        reportModel.ReadReceipt = entity.ReadReceipt == 1 ? true : false;
                        reportModel.ReadReceiptDate = entity.ReadReceiptDate;
                        reportModel.ActivityId = entity.ActivityId;
                        reportModel.DeliveryStatus = entity.DeliveryStatus;
                        reportModel.Reaction = reaction != null ? reaction.ReactionType : string.Empty;
                        reportModel.ReactionDate = reaction != null ? reaction.ReactionDate : null;
                        reportModel.UserId = entity.UserId;
                        reportModel.TenantId = entity.TenantId;

                        exportData.Add(reportModel);
                    }
                }

                await Task.Delay(0);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error in gettting GetReactionListDetail for notification id : {notificationId}");
            }

            return exportData;
        }

        /// <summary>
        /// Delete the data from the reaction table.
        /// </summary>
        /// <param name="notificationId">Notification Id.</param>
        /// <param name="activityId">Activity Id.</param>
        /// <param name="aadId">Aad Id.</param>
        /// <returns>True or False.</returns>
        public async Task<bool> DeleteReactionData(string notificationId, string activityId, string aadId)
        {
            bool isSuccess = false;
            try
            {
                var storageConnectionString = this.repositoryOptions.Value.StorageAccountConnectionString;
                CloudStorageAccount storageAccount;
                storageAccount = CloudStorageAccount.Parse(storageConnectionString);

                CloudTableClient tableClient = storageAccount.CreateCloudTableClient(new TableClientConfiguration());
                CloudTable tableUserReaction = tableClient.GetTableReference("UserReaction");

                var reactionEntities = tableUserReaction.ExecuteQuery(new TableQuery<UserReactionEntity>()).ToList().Where(x => x.PartitionKey == notificationId && x.RowKey == aadId && x.ActivityId == activityId);

                if (reactionEntities.Any())
                {
                    TableOperation reactionDeleteOperation = TableOperation.Delete(reactionEntities.First());
                    TableResult result = await tableUserReaction.ExecuteAsync(reactionDeleteOperation);
                    isSuccess = true;
                }

                await Task.Delay(0);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error in Delete Reaction Data for notification id : {notificationId} && activity id : {activityId}");
            }

            return isSuccess;
        }

        /// <summary>
        /// Delete the data from the reaction table.
        /// </summary>
        /// <param name="notificationId">Notification Id.</param>
        /// <param name="activityId">Activity Id.</param>
        /// <param name="aadId">Aad Id.</param>
        /// <returns>True or False.</returns>
        public async Task<bool> DeleteQuestionAnswerData(string notificationId, string activityId, string aadId)
        {
            bool isSuccess = false;
            try
            {
                var storageConnectionString = this.repositoryOptions.Value.StorageAccountConnectionString;
                CloudStorageAccount storageAccount;
                storageAccount = CloudStorageAccount.Parse(storageConnectionString);

                CloudTableClient tableClient = storageAccount.CreateCloudTableClient(new TableClientConfiguration());
                CloudTable tableQuestionAnswer = tableClient.GetTableReference("QuestionAnswer");

                var qnaEntities = tableQuestionAnswer.ExecuteQuery(new TableQuery<QuestionAnswerAdaptiveCardEntity>()).ToList().Where(x => x.PartitionKey == notificationId && x.RowKey == aadId && x.ActivityId == activityId);

                if (qnaEntities.Any())
                {
                    TableOperation qnaDeleteOperation = TableOperation.Delete(qnaEntities.First());
                    TableResult result = await tableQuestionAnswer.ExecuteAsync(qnaDeleteOperation);
                    isSuccess = true;
                }

                await Task.Delay(0);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error in Delete Question Answer Data for notification id : {notificationId} && activity id : {activityId}");
            }

            return isSuccess;
        }
    }
}

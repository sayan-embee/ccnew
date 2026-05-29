// <copyright file="Program.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

using System;
using System.Globalization;
using Azure.Storage.Blobs;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ExportData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RecallSentNotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.DataQueue;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.DataRecallQueue;
using Microsoft.Teams.Apps.CompanyCommunicator.Data.Func;
using Microsoft.Teams.Apps.CompanyCommunicator.Data.Func.Services.FileCardServices;
using Microsoft.Teams.Apps.CompanyCommunicator.Data.Func.Services.NotificationDataServices;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Connector.Authentication;

// Isolated worker host. Replaces FunctionsStartup-based Startup.cs.
var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureAppConfiguration(cfg =>
    {
        cfg.AddEnvironmentVariables();
    })
    .ConfigureServices((context, services) =>
    {
        var configuration = context.Configuration;

        services.AddOptions<RepositoryOptions>()
            .Configure<IConfiguration>((repositoryOptions, cfg) =>
            {
                repositoryOptions.StorageAccountConnectionString =
                    cfg.GetValue<string>("StorageAccountConnectionString");

                repositoryOptions.EnsureTableExists =
                    !cfg.GetValue<bool>("IsItExpectedThatTableAlreadyExists", true);
            });

        services.AddOptions<MessageQueueOptions>()
            .Configure<IConfiguration>((messageQueueOptions, cfg) =>
            {
                messageQueueOptions.ServiceBusConnection =
                    cfg.GetValue<string>("ServiceBusConnection");
            });

        services.AddOptions<BotOptions>()
            .Configure<IConfiguration>((botOptions, cfg) =>
            {
                botOptions.UserAppId = cfg.GetValue<string>("UserAppId");
                botOptions.UserAppPassword = cfg.GetValue<string>("UserAppPassword");
                botOptions.AuthorAppId = cfg.GetValue<string>("AuthorAppId");
                botOptions.AuthorAppPassword = cfg.GetValue<string>("AuthorAppPassword");
            });

        services.AddOptions<CleanUpFileOptions>()
            .Configure<IConfiguration>((cleanUpFileOptions, cfg) =>
            {
                cleanUpFileOptions.CleanUpFile = cfg.GetValue<string>("CleanUpFile");
            });

        services.AddOptions<DataQueueMessageOptions>()
            .Configure<IConfiguration>((dataQueueMessageOptions, cfg) =>
            {
                dataQueueMessageOptions.FirstTenMinutesRequeueMessageDelayInSeconds =
                    cfg.GetValue<double>("FirstTenMinutesRequeueMessageDelayInSeconds", 20);

                dataQueueMessageOptions.RequeueMessageDelayInSeconds =
                    cfg.GetValue<double>("RequeueMessageDelayInSeconds", 120);
            });

        services.AddLocalization();

        // Set current culture.
        var culture = Environment.GetEnvironmentVariable("i18n:DefaultCulture") ?? "en-US";
        CultureInfo.DefaultThreadCurrentCulture = new CultureInfo(culture);
        CultureInfo.DefaultThreadCurrentUICulture = new CultureInfo(culture);

        // Add blob client.
        services.AddSingleton(sp => new BlobContainerClient(
            sp.GetRequiredService<IConfiguration>().GetValue<string>("StorageAccountConnectionString"),
            Microsoft.Teams.Apps.CompanyCommunicator.Common.Constants.BlobContainerName));

        // Bot services.
        services.AddSingleton<UserAppCredentials>();
        services.AddSingleton<ICredentialProvider, ConfigurationCredentialProvider>();
        services.AddSingleton<BotFrameworkHttpAdapter>();

        // Services
        services.AddSingleton<IFileCardService, FileCardService>();
        services.AddTransient<AggregateSentNotificationDataService>();
        services.AddTransient<UpdateNotificationDataService>();
        services.AddTransient<UpdateRecallNotificationDataService>();
        services.AddTransient<AggregateRecallSentNotificationDataService>();

        // Repositories
        services.AddSingleton<INotificationDataRepository, NotificationDataRepository>();
        services.AddSingleton<ISentNotificationDataRepository, SentNotificationDataRepository>();
        services.AddSingleton<IUserDataRepository, UserDataRepository>();
        services.AddSingleton<IExportDataRepository, ExportDataRepository>();
        services.AddSingleton<IRecallSentNotificationDataRepository, RecallSentNotificationDataRepository>();

        // Service bus message queues
        services.AddSingleton<IDataQueue, DataQueue>();
        services.AddSingleton<IDataRecallQueue, DataRecallQueue>();

        // Application Insights for isolated worker.
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();
    })
    .Build();

host.Run();

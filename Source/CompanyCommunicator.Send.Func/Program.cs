// <copyright file="Program.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

using System;
using System.Globalization;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Helpers;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.RecallSentNotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.SendQueue;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.SendRecallQueue;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Teams;
using Microsoft.Teams.Apps.CompanyCommunicator.Send.Func;
using Microsoft.Teams.Apps.CompanyCommunicator.Send.Func.Services;

// Isolated worker host. Replaces FunctionsStartup-based Startup.cs.
var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureAppConfiguration(cfg =>
    {
        cfg.AddEnvironmentVariables();
    })
    .ConfigureServices((context, services) =>
    {
        services.AddOptions<SendFunctionOptions>()
            .Configure<IConfiguration>((opt, cfg) =>
            {
                opt.MaxNumberOfAttempts = cfg.GetValue<int>("MaxNumberOfAttempts", 1);
                opt.SendRetryDelayNumberOfSeconds = cfg.GetValue<double>("SendRetryDelayNumberOfSeconds", 660);
            });

        services.AddOptions<BotOptions>()
            .Configure<IConfiguration>((opt, cfg) =>
            {
                opt.UserAppId = cfg.GetValue<string>("UserAppId");
                opt.UserAppPassword = cfg.GetValue<string>("UserAppPassword");
            });

        services.AddOptions<RepositoryOptions>()
            .Configure<IConfiguration>((opt, cfg) =>
            {
                opt.StorageAccountConnectionString = cfg.GetValue<string>("StorageAccountConnectionString");
                opt.EnsureTableExists = !cfg.GetValue<bool>("IsItExpectedThatTableAlreadyExists", true);
                opt.AppBaseUrl = cfg.GetValue<string>("AppBaseUrl");
            });

        services.AddOptions<MessageQueueOptions>()
            .Configure<IConfiguration>((opt, cfg) =>
            {
                opt.ServiceBusConnection = cfg.GetValue<string>("ServiceBusConnection");
            });

        services.AddLocalization();

        var culture = Environment.GetEnvironmentVariable("i18n:DefaultCulture") ?? "en-US";
        CultureInfo.DefaultThreadCurrentCulture = new CultureInfo(culture);
        CultureInfo.DefaultThreadCurrentUICulture = new CultureInfo(culture);

        // Bot
        services.AddSingleton<UserAppCredentials>();
        services.AddSingleton<ICredentialProvider, ConfigurationCredentialProvider>();
        services.AddSingleton<BotFrameworkHttpAdapter>();

        // Teams
        services.AddTransient<IMessageService, MessageService>();

        // Repositories
        services.AddSingleton<ISendingNotificationDataRepository, SendingNotificationDataRepository>();
        services.AddSingleton<IGlobalSendingNotificationDataRepository, GlobalSendingNotificationDataRepository>();
        services.AddSingleton<ISentNotificationDataRepository, SentNotificationDataRepository>();
        services.AddSingleton<IRecallSentNotificationDataRepository, RecallSentNotificationDataRepository>();
        services.AddSingleton<INotificationDataRepository, NotificationDataRepository>();

        // Queues
        services.AddSingleton<ISendQueue, SendQueue>();
        services.AddSingleton<ISendRecallQueue, SendRecallQueue>();

        // Graph
        services.AddOptions<ConfidentialClientApplicationOptions>()
            .Configure<IConfiguration>((opt, cfg) =>
            {
                opt.ClientId = cfg.GetValue<string>("AuthorAppId");
                opt.ClientSecret = cfg.GetValue<string>("AuthorAppPassword");
                opt.TenantId = cfg.GetValue<string>("TenantId");
            });
        services.AddSingleton<IConfidentialClientApplication>(provider =>
        {
            var options = provider.GetRequiredService<IOptions<ConfidentialClientApplicationOptions>>();
            return ConfidentialClientApplicationBuilder
                .Create(options.Value.ClientId)
                .WithClientSecret(options.Value.ClientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{options.Value.TenantId}"))
                .Build();
        });
        services.AddSingleton<IAuthenticationProvider, MsalAuthenticationProvider>();
        services.AddSingleton<IGraphServiceClient>(sp => new GraphServiceClient(sp.GetRequiredService<IAuthenticationProvider>()));
        services.AddSingleton<IGraphServiceFactory, GraphServiceFactory>();
        services.AddScoped<IUsersService>(sp => sp.GetRequiredService<IGraphServiceFactory>().GetUsersService());
        services.AddScoped<IGroupMembersService>(sp => sp.GetRequiredService<IGraphServiceFactory>().GetGroupMembersService());
        services.AddScoped<IAppManagerService>(sp => sp.GetRequiredService<IGraphServiceFactory>().GetAppManagerService());
        services.AddScoped<IChatsService>(sp => sp.GetRequiredService<IGraphServiceFactory>().GetChatsService());
        services.AddSingleton<IGraphServiceFactoryByTenant, GraphServiceFactoryByTenant>();

        // Send func services
        services.AddTransient<INotificationService, NotificationService>();
        services.AddSingleton<CloudStorageHelper, CloudStorageHelper>();

        // App Insights for isolated worker
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();
    })
    .Build();

host.Run();

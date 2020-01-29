// <copyright file="Startup.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

using Microsoft.Azure.WebJobs.Hosting;
using Microsoft.Teams.App.BookAThing.SyncService;

[assembly: WebJobsStartup(typeof(Startup))]

namespace Microsoft.Teams.App.BookAThing.SyncService
{
    using System;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.Extensibility;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Identity.Client;
    using Microsoft.Teams.Apps.BookAThing.Common;
    using Microsoft.Teams.Apps.BookAThing.Common.Helpers;
    using Microsoft.Teams.Apps.BookAThing.Common.Providers.Storage;
    using Microsoft.Teams.Apps.BookAThing.SyncService.Service;

    /// <summary>
    /// Class for app configuration and injection of required dependencies.
    /// </summary>
    internal class Startup : IWebJobsStartup
    {
        /// <summary>
        /// This method gets called by the runtime. Use this method to add services to the container.
        /// </summary>
        /// <param name="builder">WebJobsBuilder service.</param>
        public void Configure(IWebJobsBuilder builder)
        {
            builder.Services.AddHttpClient();
            builder.Services.AddSingleton(new TelemetryClient(new TelemetryConfiguration(Environment.GetEnvironmentVariable("APPINSIGHTS_INSTRUMENTATIONKEY"))));
            builder.Services.AddSingleton<IFavoriteStorageProvider>(provider => new FavoriteStorageProvider(Environment.GetEnvironmentVariable("StorageConnectionString"), (TelemetryClient)provider.GetService(typeof(TelemetryClient))));
            builder.Services.AddSingleton<IRoomCollectionStorageProvider>(provider => new RoomCollectionStorageProvider(Environment.GetEnvironmentVariable("StorageConnectionString"), (TelemetryClient)provider.GetService(typeof(TelemetryClient))));
            builder.Services.AddSingleton<IExchangeSyncHelper, ExchangeSyncHelper>();
            builder.Services.AddHttpClient<IGraphApiHelper, GraphApiHelper>("GraphApiHelper", httpClient =>
            {
                httpClient.BaseAddress = new Uri("https://graph.microsoft.com");
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
            });

            builder.Services.AddSingleton<IGraphApiHelper, GraphApiHelper>();
            builder.Services.AddSingleton<ISearchService>(provider => new SearchService(
                Environment.GetEnvironmentVariable("SearchServiceName"),
                Environment.GetEnvironmentVariable("SearchServiceAdminApiKey"),
                Environment.GetEnvironmentVariable("SearchIndexingIntervalInMinutes"),
                Environment.GetEnvironmentVariable("StorageConnectionString"),
                null,
                (TelemetryClient)provider.GetService(typeof(TelemetryClient))));

            builder.Services.AddSingleton(ConfidentialClientApplicationBuilder.Create(Environment.GetEnvironmentVariable("ClientId"))
                .WithAuthority(AzureCloudInstance.AzurePublic, Environment.GetEnvironmentVariable("TenantId"))
                .WithClientSecret(Environment.GetEnvironmentVariable("ClientSecret")).Build());
        }
    }
}

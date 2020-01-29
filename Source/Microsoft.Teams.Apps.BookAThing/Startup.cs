// <copyright file="Startup.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing
{
    using System;
    using System.Collections.Generic;
    using System.Net.Http;
    using System.Text;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authentication.JwtBearer;
    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Azure;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.IdentityModel.Tokens;
    using Microsoft.Teams.Apps.BookAThing.Bots;
    using Microsoft.Teams.Apps.BookAThing.Common;
    using Microsoft.Teams.Apps.BookAThing.Common.Helpers;
    using Microsoft.Teams.Apps.BookAThing.Common.Providers;
    using Microsoft.Teams.Apps.BookAThing.Common.Providers.Storage;
    using Microsoft.Teams.Apps.BookAThing.Dialogs;
    using Microsoft.Teams.Apps.BookAThing.Helpers;
    using Microsoft.Teams.Apps.BookAThing.Providers.Storage;
    using Newtonsoft.Json.Serialization;
    using Polly;
    using Polly.Extensions.Http;

    /// <summary>
    /// Class for app configuration and injection of required dependencies.
    /// </summary>
    public class Startup
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Startup"/> class.
        /// </summary>
        /// <param name="configuration">Configuration settings.</param>
        public Startup(Extensions.Configuration.IConfiguration configuration)
        {
            this.Configuration = configuration;
        }

        /// <summary>
        /// Gets Configurations Interfaces.
        /// </summary>
        public Extensions.Configuration.IConfiguration Configuration { get; }

        /// <summary>
        /// This method gets called by the runtime. Use this method to add services to the container.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddHttpClient<IGraphApiHelper, GraphApiHelper>("GraphApiHelper", httpClient =>
            {
                httpClient.BaseAddress = new Uri("https://graph.microsoft.com");
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
            }).AddPolicyHandler(GetRetryPolicy());

            services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme).AddJwtBearer(options =>
            {
                options.TokenValidationParameters = new TokenValidationParameters
                {
                    ValidateAudience = true,
                    ValidAudiences = new List<string> { this.Configuration["AppBaseUri"] },
                    ValidIssuers = new List<string> { this.Configuration["AppBaseUri"] },
                    ValidateIssuer = true,
                    ValidateIssuerSigningKey = true,
                    IssuerSigningKey = new SymmetricSecurityKey(Encoding.ASCII.GetBytes(this.Configuration["SecurityKey"])),
                    RequireExpirationTime = true,
                    ValidateLifetime = true,
                    ClockSkew = TimeSpan.FromMinutes(5),
                };
            });

            services.AddSingleton(new OAuthClient(new MicrosoftAppCredentials(this.Configuration["MicrosoftAppId"], this.Configuration["MicrosoftAppPassword"])));
            services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_1).AddJsonOptions(options => options.SerializerSettings.ContractResolver = new DefaultContractResolver());
            services.AddSingleton<ITokenHelper>(provider => new TokenHelper(this.Configuration["SecurityKey"], this.Configuration["AppBaseUri"], this.Configuration["ConnectionName"], (OAuthClient)provider.GetService(typeof(OAuthClient))));

            // Create the Bot Framework Adapter with error handling enabled.
            services.AddSingleton<IBotFrameworkHttpAdapter, AdapterWithErrorHandler>();

            // For conversation state.
            services.AddSingleton<IStorage>(new AzureBlobStorage(this.Configuration["StorageConnectionString"], "bot-state"));

            // Create the User state. (Used in this bot's Dialog implementation.)
            services.AddSingleton<UserState>();
            services.AddSingleton<IActivityStorageProvider>(provider => new ActivityStorageProvider(this.Configuration["StorageConnectionString"], (TelemetryClient)provider.GetService(typeof(TelemetryClient))));
            services.AddSingleton<IFavoriteStorageProvider>(provider => new FavoriteStorageProvider(this.Configuration["StorageConnectionString"], (TelemetryClient)provider.GetService(typeof(TelemetryClient))));
            services.AddSingleton<IRoomCollectionStorageProvider>(provider => new RoomCollectionStorageProvider(this.Configuration["StorageConnectionString"], (TelemetryClient)provider.GetService(typeof(TelemetryClient))));
            services.AddSingleton<IUserConfigurationStorageProvider>(provider => new UserConfigurationStorageProvider(this.Configuration["StorageConnectionString"], (TelemetryClient)provider.GetService(typeof(TelemetryClient))));
            services.AddSingleton<ISearchService>(provider => new SearchService(this.Configuration["SearchServiceName"], this.Configuration["SearchServiceAdminApiKey"], this.Configuration["SearchIndexingIntervalInMinutes"], this.Configuration["StorageConnectionString"], this.Configuration["SearchServiceQueryApiKey"], (TelemetryClient)provider.GetService(typeof(TelemetryClient))));
            services.AddSingleton<IMeetingHelper, MeetingHelper>();
            services.AddSingleton<IGraphApiHelper, GraphApiHelper>();
            services.AddSingleton(new MicrosoftAppCredentials(this.Configuration["MicrosoftAppId"], this.Configuration["MicrosoftAppPassword"]));

            // Create the Conversation state. (Used by the Dialog system itself.)
            services.AddSingleton<ConversationState>();
            services.AddSingleton<TelemetryClient>();
            services.AddSingleton<IMeetingProvider, MeetingProvider>();
            services.AddSingleton<IUserConfigurationProvider, UserConfigurationProvider>();

            // The Dialog that will be run by the bot.
            services.AddSingleton<MainDialog>();

            // Create the bot as a transient. In this case the ASP Controller is expecting an IBot.
            services.AddTransient<IBot>(provider => new BookAMeetingBot<MainDialog>(
                (ConversationState)provider.GetService(typeof(ConversationState)),
                (UserState)provider.GetService(typeof(UserState)),
                (MainDialog)provider.GetService(typeof(MainDialog)),
                (ITokenHelper)provider.GetService(typeof(ITokenHelper)),
                (IActivityStorageProvider)provider.GetService(typeof(IActivityStorageProvider)),
                (IFavoriteStorageProvider)provider.GetService(typeof(IFavoriteStorageProvider)),
                (IMeetingProvider)provider.GetService(typeof(IMeetingProvider)),
                (TelemetryClient)provider.GetService(typeof(TelemetryClient)),
                (IUserConfigurationStorageProvider)provider.GetService(typeof(IUserConfigurationStorageProvider)),
                this.Configuration["AppBaseUri"],
                this.Configuration["APPINSIGHTS_INSTRUMENTATIONKEY"],
                this.Configuration["TenantId"]));
        }

        /// <summary>
        /// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        /// </summary>
        /// <param name="app">Provides the mechanisms to configure an application's request pipeline.</param>
        /// <param name="env">Provides application-management functions and application services to a managed application within its application domain.</param>
        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseHsts();
            }

            app.UseAuthentication();
            app.UseDefaultFiles();
            app.UseStaticFiles();
            app.UseMvc();
        }

        /// <summary>
        /// Retry policy with jitter. Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </summary>
        /// <returns>Policy.</returns>
        private static IAsyncPolicy<HttpResponseMessage> GetRetryPolicy()
        {
            return HttpPolicyExtensions
                .HandleTransientHttpError()
                .OrResult(response => response.IsSuccessStatusCode == false)
                .WaitAndRetryAsync(2, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)));
        }
    }
}

// <copyright file="AdapterWithErrorHandler.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing
{
    using System;
    using Microsoft.ApplicationInsights;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.BookAThing.Resources;

    /// <summary>
    /// Class to handle errors and exception occured in bot.
    /// </summary>
    public class AdapterWithErrorHandler : BotFrameworkHttpAdapter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AdapterWithErrorHandler"/> class.
        /// </summary>
        /// <param name="configuration">Application configuration.</param>
        /// <param name="telemetryClient">Telemetry client for logging events and errors.</param>
        /// <param name="conversationState">Reads and writes conversation state for your bot to storage.</param>
        public AdapterWithErrorHandler(IConfiguration configuration, TelemetryClient telemetryClient, ConversationState conversationState = null)
            : base(configuration)
        {
            this.OnTurnError = async (turnContext, exception) =>
            {
                // Log any leaked exception from the application.
                telemetryClient.TrackException(exception);

                // Send a catch-all apology to the user.
                await turnContext.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                if (conversationState != null)
                {
                    try
                    {
                        // Delete the conversationState for the current conversation to prevent the
                        // bot from getting stuck in a error-loop caused by being in a bad state.
                        // ConversationState should be thought of as similar to "cookie-state" in a Web pages.
                        await conversationState.DeleteAsync(turnContext).ConfigureAwait(false);
                    }
                    catch (Exception e)
                    {
                        telemetryClient.TrackException(e);
                    }
                }
            };
        }
    }
}

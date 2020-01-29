// <copyright file="LogoutDialog.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Dialogs
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Dialogs;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.BookAThing.Resources;

    /// <summary>
    /// Dialog for handling interruption.
    /// </summary>
    public class LogoutDialog : ComponentDialog
    {
        /// <summary>
        /// Telemetry client to log events and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="LogoutDialog"/> class.
        /// </summary>
        /// <param name="id">Dialog Id.</param>
        /// <param name="connectionName">Connection name for Active Directory authentication.</param>
        /// <param name="telemetryClient">Telemetry client to log events and errors.</param>
        public LogoutDialog(string id, string connectionName, TelemetryClient telemetryClient)
            : base(id)
        {
            this.ConnectionName = connectionName;
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Gets connection name for Active Directory authentication.
        /// </summary>
        protected string ConnectionName { get; private set; }

        /// <summary>
        /// Overriding OnBeginDialogAsync method for checking interruption.
        /// </summary>
        /// <param name="innerDialogContext">Child dialog context.</param>
        /// <param name="options">Options object.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task<DialogTurnResult> OnBeginDialogAsync(DialogContext innerDialogContext, object options, CancellationToken cancellationToken)
        {
            var dialogTurnContext = await this.OnInterruptAsync(innerDialogContext, cancellationToken).ConfigureAwait(false);
            if (dialogTurnContext != null)
            {
                return dialogTurnContext;
            }

            return await base.OnBeginDialogAsync(innerDialogContext, options, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Overriding OnContinueDialogAsync method for checking interruption.
        /// </summary>
        /// <param name="innerDialogContext">Child dialog context.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task<DialogTurnResult> OnContinueDialogAsync(DialogContext innerDialogContext, CancellationToken cancellationToken)
        {
            var dialogTurnContext = await this.OnInterruptAsync(innerDialogContext, cancellationToken).ConfigureAwait(false);
            if (dialogTurnContext != null)
            {
                return dialogTurnContext;
            }

            return await base.OnContinueDialogAsync(innerDialogContext, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Handling interruption.
        /// </summary>
        /// <param name="innerDialogContext">Child dialog context.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task<DialogTurnResult> OnInterruptAsync(DialogContext innerDialogContext, CancellationToken cancellationToken)
        {
            if (innerDialogContext.Context.Activity.Type == ActivityTypes.Message)
            {
                var text = innerDialogContext.Context.Activity.Text;
                var botAdapter = (BotFrameworkAdapter)innerDialogContext.Context.Adapter;

                if (!string.IsNullOrEmpty(text))
                {
                    if (text.Equals(BotCommands.Logout, StringComparison.OrdinalIgnoreCase))
                    {
                        this.telemetryClient.TrackTrace(innerDialogContext.Context.Activity.From.AadObjectId + " typed " + text);
                        await botAdapter.SignOutUserAsync(innerDialogContext.Context, this.ConnectionName, null, cancellationToken).ConfigureAwait(false);

                        await innerDialogContext.Context.SendActivityAsync(MessageFactory.Text(Strings.LogoutSuccess), cancellationToken).ConfigureAwait(false);
                        return await innerDialogContext.CancelAllDialogsAsync().ConfigureAwait(false);
                    }
                }
            }

            return null;
        }
    }
}

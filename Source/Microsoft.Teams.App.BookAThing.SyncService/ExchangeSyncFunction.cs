// <copyright file="ExchangeSyncFunction.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.BookAThing.SyncService
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.BookAThing.SyncService.Service;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// Exchange sync function to perform rooms sync with Azure Table Storage.
    /// </summary>
    public class ExchangeSyncFunction
    {
        /// <summary>
        /// Retry policy with jitter, Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </summary>
        private static RetryPolicy retryPolicy = Policy.Handle<Exception>().WaitAndRetryAsync(Backoff.DecorrelatedJitterBackoffV2(TimeSpan.FromMilliseconds(1000), 2));

        /// <summary>
        /// Exchange sync helper.
        /// </summary>
        private readonly IExchangeSyncHelper exchangeSyncHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeSyncFunction"/> class.
        /// </summary>
        /// <param name="exchangeSyncHelper">Exchange sync helper.</param>
        public ExchangeSyncFunction(IExchangeSyncHelper exchangeSyncHelper)
        {
            this.exchangeSyncHelper = exchangeSyncHelper;
        }

        /// <summary>
        /// Runs trigger for every sunday at 12:00 AM (default schedule expression "0 0 0 * * 0").
        /// </summary>
        /// <param name="timerInfo">Timer info object.</param>
        /// <param name="log">Trace writer object for logging.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [FunctionName("ExchangeSyncFunction")]
        public async Task Run([TimerTrigger("0 0 0 * * 0", RunOnStartup = true)]TimerInfo timerInfo, ILogger log)
        {
            log.LogInformation("Exchange sync azure function started");
            await retryPolicy.ExecuteAsync(() => this.exchangeSyncHelper.ExchangeToStorageExportAsync()).ConfigureAwait(false);
        }
    }
}

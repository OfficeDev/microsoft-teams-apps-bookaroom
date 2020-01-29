// <copyright file="BotController.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Controllers
{
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;

    /// <summary>
    /// This ASP Controller is created to handle a request. Dependency Injection will provide the Adapter and IBot implementation at runtime.
    /// Multiple different IBot implementations running at different endpoints can be
    /// achieved by specifying a more specific type for the bot constructor argument.
    /// </summary>
    [Route("api/messages")]
    [ApiController]
    public class BotController : ControllerBase
    {
        /// <summary>
        /// Bot adapter.
        /// </summary>
        private readonly IBotFrameworkHttpAdapter adapter;

        /// <summary>
        /// Book A Room bot.
        /// </summary>
        private readonly IBot bot;

        /// <summary>
        /// Initializes a new instance of the <see cref="BotController"/> class.
        /// </summary>
        /// <param name="adapter">A Bot Builder Adapter implementation used to handled bot Framework HTTP requests.</param>
        /// <param name="bot">Book a Room bot which contains core logic.</param>
        public BotController(IBotFrameworkHttpAdapter adapter, IBot bot)
        {
            this.adapter = adapter;
            this.bot = bot;
        }

        /// <summary>
        /// Forwards post request to bot.
        /// </summary>
        /// <returns>A <see cref="Task"/> representing the result of the asynchronous operation.</returns>
        [HttpPost]
        public async Task PostAsync()
        {
            // Delegate the processing of the HTTP POST to the adapter.
            // The adapter will invoke the bot.
          await this.adapter.ProcessAsync(this.Request, this.Response, this.bot).ConfigureAwait(false);
        }
    }
}

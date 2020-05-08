// <copyright file="AllowedCallersClaimsValidator.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Authentication
{
    using System;
    using System.Collections.Generic;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;

    /// <summary>
    /// Sample claims validator that loads an allowed list from configuration if present
    /// and checks that requests are coming from allowed parent bots.
    /// </summary>
    public class AllowedCallersClaimsValidator : ClaimsValidator
    {
        private const string ConfigKey = "AllowedCallers";
        private readonly List<string> allowedCallers;

        /// <summary>
        /// Initializes a new instance of the <see cref="AllowedCallersClaimsValidator"/> class.
        /// </summary>
        /// <param name="configuration">Application configuration.</param>
        public AllowedCallersClaimsValidator(IConfiguration configuration)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException(nameof(configuration));
            }

            // AllowedCallers is the setting in the appsettings.json file
            // that consists of the list of parent bot IDs that are allowed to access the skill.
            // To add a new parent bot, simply edit the AllowedCallers and add
            // the parent bot's Microsoft app ID to the list.
            // In this sample, we allow all callers if AllowedCallers contains an "*".
            var section = configuration.GetSection(ConfigKey);
            var appsList = section.Get<string[]>();
            if (appsList == null)
            {
                throw new ArgumentNullException($"\"{ConfigKey}\" not found in configuration.");
            }

            this.allowedCallers = new List<string>(appsList);
        }

        /// <summary>
        /// Validates parent caller.
        /// </summary>
        /// <param name="claims">The list of claims to validate.</param>
        /// <returns> The completed task if validation is successful or throws UnauthorizedAccessException.</returns>
        public override Task ValidateClaimsAsync(IList<Claim> claims)
        {
            // If _allowedCallers contains an "*", we allow all callers.
            if (SkillValidation.IsSkillClaim(claims) && !this.allowedCallers.Contains("*"))
            {
                // Check that the appId claim in the skill request is in the list of callers configured for this bot.
                var appId = JwtTokenValidation.GetAppIdFromClaims(claims);

                if (!this.allowedCallers.Contains(appId))
                {
                    throw new UnauthorizedAccessException($"Received a request from an application with an appID of \"{appId}\". To enable requests from this skill, add the skill to your configuration file.");
                }
            }

            return Task.CompletedTask;
        }
    }
}

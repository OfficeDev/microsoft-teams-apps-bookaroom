// <copyright file="TokenHelper.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.IdentityModel.Tokens.Jwt;
    using System.Security.Claims;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.Bot.Connector;
    using Microsoft.IdentityModel.Tokens;
    using Microsoft.Teams.Apps.BookAThing.Common;

    /// <summary>
    /// Helper class for JWT token generation and validation.
    /// </summary>
    public class TokenHelper : ITokenHelper
    {
        /// <summary>
        /// Used to retrieve user Active Directory access token from Bot Framework.
        /// </summary>
        private readonly OAuthClient oAuthClient;

        /// <summary>
        /// Security key for generating and validating token.
        /// </summary>
        private readonly string securityKey;

        /// <summary>
        /// Application base Url.
        /// </summary>
        private readonly string appBaseUri;

        /// <summary>
        /// AAD authentication connection name.
        /// </summary>
        private readonly string connectionName;

        /// <summary>
        /// Telemetry client to log event and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="TokenHelper"/> class.
        /// </summary>
        /// <param name="securityKey">Security key for generating and validating token.</param>
        /// <param name="appBaseUri">Application base Url.</param>
        /// <param name="connectionName">Active Directory authentication connection name.</param>
        /// <param name="oAuthClient">Used to retrieve user Active Directory access token from Bot Framework.</param>
        /// <param name="telemetryClient">Telemetry client to log event and errors.</param>
        public TokenHelper(string securityKey, string appBaseUri, string connectionName, OAuthClient oAuthClient, TelemetryClient telemetryClient)
        {
            this.securityKey = securityKey;
            this.appBaseUri = appBaseUri;
            this.connectionName = connectionName;
            this.oAuthClient = oAuthClient;
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Generate JWT token used by client app to authenticate HTTP calls with API.
        /// </summary>
        /// <param name="userObjectIdentifer">User AD object id.</param>
        /// <param name="serviceURL">Service URL from bot.</param>
        /// <param name="fromId">Unique Id from activity.</param>
        /// <param name="jwtExpiryMinutes">Expiry of token.</param>
        /// <returns>JWT token.</returns>
        public string GenerateAPIAuthToken(string userObjectIdentifer, string serviceURL, string fromId, int jwtExpiryMinutes)
        {
            SymmetricSecurityKey signingKey = new SymmetricSecurityKey(Encoding.ASCII.GetBytes(this.securityKey));
            SigningCredentials signingCredentials = new SigningCredentials(signingKey, SecurityAlgorithms.HmacSha256);

            SecurityTokenDescriptor securityTokenDescriptor = new SecurityTokenDescriptor()
            {
                Subject = new ClaimsIdentity(
                    new List<Claim>()
                    {
                        new Claim("userObjectIdentifer", userObjectIdentifer),
                        new Claim("serviceURL", serviceURL),
                        new Claim("fromId", fromId),
                    }, "Custom"),
                NotBefore = DateTime.UtcNow,
                SigningCredentials = signingCredentials,
                Issuer = this.appBaseUri,
                Audience = this.appBaseUri,
                IssuedAt = DateTime.UtcNow,
                Expires = DateTime.UtcNow.AddMinutes(jwtExpiryMinutes),
            };

            JwtSecurityTokenHandler tokenHandler = new JwtSecurityTokenHandler();
            SecurityToken token = tokenHandler.CreateToken(securityTokenDescriptor);
            return tokenHandler.WriteToken(token);
        }

        /// <summary>
        /// Get Active Directory access token for user.
        /// </summary>
        /// <param name="fromId">Activity.From.Id from bot.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task<string> GetUserTokenAsync(string fromId)
        {
            try
            {
                var token = await this.oAuthClient.UserToken.GetAadTokensAsync(fromId, this.connectionName, new Bot.Schema.AadResourceUrls { ResourceUrls = new string[] { Constants.GraphAPIBaseUrl } }).ConfigureAwait(false);
                return token?[Constants.GraphAPIBaseUrl]?.Token;
            }
            catch (Exception ex)
            {
                // scenarios that will throw execptions are:
                // 1. ValidationException: properties passed to GetAadTokensAsync are invalid  or null
                // 2. bot service failed to fetch user token and throws exception "Operation returned an invalid status code"
                this.telemetryClient.TrackException(ex);
                return null;
            }
        }
    }
}

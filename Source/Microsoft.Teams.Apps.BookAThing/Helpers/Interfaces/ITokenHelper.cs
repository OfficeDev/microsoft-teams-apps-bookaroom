// <copyright file="ITokenHelper.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Helpers
{
    using System.Threading.Tasks;

    /// <summary>
    /// Helper for JWT token generation and validation.
    /// </summary>
    public interface ITokenHelper
    {
        /// <summary>
        /// Generate JWT token used by client app to authenticate HTTP calls with API.
        /// </summary>
        /// <param name="userObjectIdentifer">User Active Directory object id.</param>
        /// <param name="serviceURL">Service URL from bot.</param>
        /// <param name="fromId">Unique Id from activity.</param>
        /// <param name="jwtExpiryMinutes">Expiry of token.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        string GenerateAPIAuthToken(string userObjectIdentifer, string serviceURL, string fromId, int jwtExpiryMinutes);

        /// <summary>
        /// Get Active Directory access token for user.
        /// </summary>
        /// <param name="fromId">Activity.From.Id from bot.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<string> GetUserTokenAsync(string fromId);
    }
}
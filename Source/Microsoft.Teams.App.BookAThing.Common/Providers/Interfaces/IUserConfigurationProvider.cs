// <copyright file="IUserConfigurationProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Providers
{
    using System.Threading.Tasks;
    using Microsoft.Teams.App.BookAThing.Common.Models.Response;

    /// <summary>
    /// Exposes methods for fetching user specific data.
    /// </summary>
    public interface IUserConfigurationProvider
    {
        /// <summary>
        /// Get supported time zones for signed in user.
        /// </summary>
        /// <param name="token">Active Directory access token.</param>
        /// <returns>List of supported time zones.</returns>
        Task<SupportedTimeZoneResponse> GetSupportedTimeZoneAsync(string token);
    }
}
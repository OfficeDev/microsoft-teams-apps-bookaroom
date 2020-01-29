// <copyright file="TimeZoneResponse.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.BookAThing.Common.Models.Response
{
    using Newtonsoft.Json;

    /// <summary>
    /// Supported time zone response from Graph API.
    /// </summary>
    public class TimeZoneResponse
    {
        /// <summary>
        /// Gets or sets alias for time zone.
        /// </summary>
        [JsonProperty("alias")]
        public string Alias { get; set; }

        /// <summary>
        /// Gets or sets display name for time zone.
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }
    }
}

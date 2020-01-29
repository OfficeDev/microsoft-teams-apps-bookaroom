// <copyright file="Location.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Models.Request
{
    using Newtonsoft.Json;

    /// <summary>
    /// Location class.
    /// </summary>
    public class Location
    {
        /// <summary>
        /// Gets or sets display name of location.
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }
    }
}

// <copyright file="PlaceInfo.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.BookAThing.Common.Models.Response
{
    using Newtonsoft.Json;

    /// <summary>
    /// Class describing place information.
    /// </summary>
    public class PlaceInfo
    {
        /// <summary>
        /// Gets or sets id of place.
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets name of place.
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        /// <summary>
        /// Gets or sets email address associated with place.
        /// </summary>
        [JsonProperty("emailAddress")]
        public string EmailAddress { get; set; }
    }
}

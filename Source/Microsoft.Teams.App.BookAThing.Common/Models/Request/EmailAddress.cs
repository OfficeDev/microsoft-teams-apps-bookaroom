// <copyright file="EmailAddress.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Models.Request
{
    using Newtonsoft.Json;

    /// <summary>
    /// Email address and display name of room/building.
    /// </summary>
    public class EmailAddress
    {
        /// <summary>
        /// Gets or sets email address.
        /// </summary>
        [JsonProperty("address")]
        public string Address { get; set; }

        /// <summary>
        /// Gets or sets name.
        /// </summary>
        [JsonProperty("name")]
        public string Name { get; set; }
    }
}

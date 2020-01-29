// <copyright file="Body.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Models.Request
{
    using Newtonsoft.Json;

    /// <summary>
    /// Used as request body for CreateEvent Graph API call.
    /// </summary>
    public class Body
    {
        /// <summary>
        /// Gets or sets content type for request.
        /// </summary>
        [JsonProperty("contentType")]
        public string ContentType { get; set; }

        /// <summary>
        /// Gets or sets content for request.
        /// </summary>
        [JsonProperty("content")]
        public string Content { get; set; }
    }
}

// <copyright file="SupportedTimeZoneResponse.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.BookAThing.Common.Models.Response
{
    using System.Collections.Generic;
    using System.Net;
    using Microsoft.Teams.App.BookAThing.Common.Models.Error;
    using Newtonsoft.Json;

    /// <summary>
    /// Supported time zones from Graph API.
    /// </summary>
    public class SupportedTimeZoneResponse
    {
        /// <summary>
        /// Gets or sets list of time zones.
        /// </summary>
        [JsonProperty("value")]
        public List<TimeZoneResponse> TimeZones { get; set; }

        /// <summary>
        /// Gets or sets Graph API response status code.
        /// </summary>
        public HttpStatusCode StatusCode { get; set; }

        /// <summary>
        /// Gets or sets Graph API error response.
        /// </summary>
        public ErrorResponse ErrorResponse { get; set; }
    }
}

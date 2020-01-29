// <copyright file="PlaceResponse.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.BookAThing.Common.Models.Response
{
    using System.Collections.Generic;
    using Microsoft.Teams.App.BookAThing.Common.Models.Error;
    using Newtonsoft.Json;

    /// <summary>
    /// Places response class.
    /// </summary>
    public class PlaceResponse
    {
        /// <summary>
        /// Gets or sets list of places.
        /// </summary>
        [JsonProperty("value")]
        public List<PlaceInfo> PlaceDetails { get; set; }

        /// <summary>
        /// Gets or sets Graph API error response.
        /// </summary>
        public ErrorResponse ErrorResponse { get; set; }
    }
}

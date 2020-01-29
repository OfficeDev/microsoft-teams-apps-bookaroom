// <copyright file="ErrorResponse.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.BookAThing.Common.Models.Error
{
    using Newtonsoft.Json;

    /// <summary>
    /// Error response root object for Graph API.
    /// </summary>
    public class ErrorResponse
    {
        /// <summary>
        /// Gets or sets error root object.
        /// </summary>
        [JsonProperty("error")]
        public Error Error { get; set; }
    }
}

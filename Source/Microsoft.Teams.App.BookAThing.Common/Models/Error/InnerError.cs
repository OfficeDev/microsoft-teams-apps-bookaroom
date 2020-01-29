// <copyright file="InnerError.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.BookAThing.Common.Models.Error
{
    using System;
    using Newtonsoft.Json;

    /// <summary>
    /// Inner error class.
    /// </summary>
    public class InnerError
    {
        /// <summary>
        /// Gets or sets Graph API request ID.
        /// </summary>
        [JsonProperty("request-id")]
        public string RequestId { get; set; }

        /// <summary>
        /// Gets or sets date time of error response.
        /// </summary>
        [JsonProperty("date")]
        public DateTime Date { get; set; }
    }
}

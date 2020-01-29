// <copyright file="ResponseStatus.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Models.Response
{
    using System;
    using Newtonsoft.Json;

    /// <summary>
    /// Response status class.
    /// </summary>
    public class ResponseStatus
    {
        /// <summary>
        /// Gets or sets response.
        /// </summary>
        [JsonProperty("response")]
        public string Response { get; set; }

        /// <summary>
        /// Gets or sets time.
        /// </summary>
        [JsonProperty("time")]
        public DateTime Time { get; set; }
    }
}

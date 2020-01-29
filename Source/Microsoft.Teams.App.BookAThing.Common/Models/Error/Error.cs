// <copyright file="Error.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.App.BookAThing.Common.Models.Error
{
    using Newtonsoft.Json;

    /// <summary>
    /// Error response class.
    /// </summary>
    public class Error
    {
        /// <summary>
        /// Gets or sets error status code.
        /// </summary>
        [JsonProperty("code")]
        public string StatusCode { get; set; }

        /// <summary>
        /// Gets or sets error message.
        /// </summary>
        [JsonProperty("message")]
        public string ErrorMessage { get; set; }

        /// <summary>
        /// Gets or sets inner error information.
        /// </summary>
        [JsonProperty("innerError")]
        public InnerError InnerError { get; set; }
    }
}

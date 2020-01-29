// <copyright file="CreateEventResponse.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Models.Response
{
    using System.Net;
    using Microsoft.Teams.App.BookAThing.Common.Models.Error;
    using Microsoft.Teams.Apps.BookAThing.Common.Models;
    using Newtonsoft.Json;

    /// <summary>
    /// Response for meeting creation.
    /// </summary>
    public class CreateEventResponse
    {
        /// <summary>
        /// Gets or sets id of meeting.
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets date for meeting creation.
        /// </summary>
        [JsonProperty("createdDateTime")]
        public System.DateTime CreatedDateTime { get; set; }

        /// <summary>
        /// Gets or sets date for meeting modification.
        /// </summary>
        [JsonProperty("lastModifiedDateTime")]
        public System.DateTime LastModifiedDateTime { get; set; }

        /// <summary>
        /// Gets or sets subject for meeting.
        /// </summary>
        [JsonProperty("subject")]
        public string Subject { get; set; }

        /// <summary>
        /// Gets or sets type of meeting.
        /// </summary>
        [JsonProperty("type")]
        public string Type { get; set; }

        /// <summary>
        /// Gets or sets web link for meeting.
        /// </summary>
        [JsonProperty("webLink")]
        public string WebLink { get; set; }

        /// <summary>
        /// Gets or sets online meeting url.
        /// </summary>
        [JsonProperty("onlineMeetingUrl")]
        public object OnlineMeetingUrl { get; set; }

        /// <summary>
        /// Gets or sets response status.
        /// </summary>
        [JsonProperty("responseStatus")]
        public ResponseStatus ResponseStatus { get; set; }

        /// <summary>
        /// Gets or sets meeting start time with time zone.
        /// </summary>
        [JsonProperty("start")]
        public DateTimeAndTimeZone Start { get; set; }

        /// <summary>
        /// Gets or sets meeting end time with time zone.
        /// </summary>
        [JsonProperty("end")]
        public DateTimeAndTimeZone End { get; set; }

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

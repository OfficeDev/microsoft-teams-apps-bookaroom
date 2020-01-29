// <copyright file="CreateEventRequest.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Models.Request
{
    using System.Collections.Generic;
    using Microsoft.Teams.Apps.BookAThing.Common.Models;
    using Newtonsoft.Json;

    /// <summary>
    /// Model used as body for sending meeting creation request to Graph API.
    /// </summary>
    public class CreateEventRequest
    {
        /// <summary>
        /// Gets or sets subject of meeting.
        /// </summary>
        [JsonProperty("subject")]
        public string Subject { get; set; }

        /// <summary>
        /// Gets or sets body for request.
        /// </summary>
        [JsonProperty("body")]
        public Body Body { get; set; }

        /// <summary>
        /// Gets or sets start time with time zone.
        /// </summary>
        [JsonProperty("start")]
        public DateTimeAndTimeZone Start { get; set; }

        /// <summary>
        /// Gets or sets end time with time zone.
        /// </summary>
        [JsonProperty("end")]
        public DateTimeAndTimeZone End { get; set; }

        /// <summary>
        /// Gets or sets location of meeting.
        /// </summary>
        [JsonProperty("location")]
        public Location Location { get; set; }

        /// <summary>
        /// Gets or sets attendees list.
        /// </summary>
        [JsonProperty("attendees")]
        public List<Attendee> Attendees { get; set; }
    }
}

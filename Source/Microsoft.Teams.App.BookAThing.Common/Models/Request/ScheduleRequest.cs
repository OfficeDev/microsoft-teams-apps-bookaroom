// <copyright file="ScheduleRequest.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Models.Request
{
    using System.Collections.Generic;
    using Newtonsoft.Json;

    /// <summary>
    /// Body for sending schedule request to Graph API.
    /// </summary>
    public class ScheduleRequest
    {
        /// <summary>
        /// Gets or sets list of rooms for which schedule needs to be fetched.
        /// </summary>
        [JsonProperty("schedules")]
        public List<string> Schedules { get; set; }

        /// <summary>
        /// Gets or sets start time with time zone.
        /// </summary>
        [JsonProperty("startTime")]
        public DateTimeAndTimeZone StartDateTime { get; set; }

        /// <summary>
        /// Gets or sets end time with time zone.
        /// </summary>
        [JsonProperty("endTime")]
        public DateTimeAndTimeZone EndDateTime { get; set; }
    }
}

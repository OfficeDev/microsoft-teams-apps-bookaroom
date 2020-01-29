// <copyright file="ScheduleItem.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Models.Response
{
    using Microsoft.Teams.Apps.BookAThing.Common.Models;
    using Newtonsoft.Json;

    /// <summary>
    /// Schedule details for room.
    /// </summary>
    public class ScheduleItem
    {
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
    }
}

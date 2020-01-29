// <copyright file="Schedule.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Models.Response
{
    using System.Collections.Generic;
    using Newtonsoft.Json;

    /// <summary>
    /// Schedule meeting room.
    /// </summary>
    public class Schedule
    {
        /// <summary>
        /// Gets or sets schedule Id.
        /// </summary>
        [JsonProperty("scheduleId")]
        public string ScheduleId { get; set; }

        /// <summary>
        /// Gets or sets room name.
        /// </summary>
        [JsonProperty("roomName")]
        public string RoomName { get; set; }

        /// <summary>
        /// Gets or sets building name.
        /// </summary>
        [JsonProperty("buildingName")]
        public string BuildingName { get; set; }

        /// <summary>
        /// Gets or sets room availability.
        /// </summary>
        [JsonProperty("availabilityView")]
        public string AvailabilityView { get; set; }

        /// <summary>
        /// Gets or sets schedule items. Each item represents schedule for given timespan.
        /// </summary>
        [JsonProperty("scheduleItems")]
        public List<ScheduleItem> ScheduleItems { get; set; }
    }
}

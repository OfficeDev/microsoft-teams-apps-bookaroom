// <copyright file="Meeting.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Class containing meeting details.
    /// </summary>
    public class Meeting
    {
        /// <summary>
        /// Gets or sets email associated with room.
        /// </summary>
        public string RoomEmail { get; set; }

        /// <summary>
        /// Gets or sets name of room.
        /// </summary>
        public string RoomName { get; set; }

        /// <summary>
        /// Gets or sets name of building.
        /// </summary>
        public string BuildingName { get; set; }

        /// <summary>
        /// Gets or sets start time for meeting.
        /// </summary>
        public string StartDateTime { get; set; }

        /// <summary>
        /// Gets or sets end time for meeting.
        /// </summary>
        public string EndDateTime { get; set; }

        /// <summary>
        /// Gets or sets status.
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// Gets or sets text.
        /// </summary>
        [JsonProperty("text")]
        public string Text { get; set; }
    }
}

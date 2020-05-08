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
        /// Initializes a new instance of the <see cref="Meeting"/> class.
        /// </summary>
        /// <param name="skillId"> Microsoft app id to embed in card actions.</param>
        public Meeting(string skillId)
        {
            this.SkillId = skillId;
        }

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

        /// <summary>
        /// Gets or sets skillId.
        /// </summary>
        [JsonProperty("skillId")]
        public string SkillId { get; set; }
    }
}

// <copyright file="CreateEvent.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Models
{
    /// <summary>
    /// Class containing required properties to create meeting.
    /// </summary>
    public class CreateEvent
    {
        /// <summary>
        /// Gets or sets name of building.
        /// </summary>
        public string BuildingName { get; set; }

        /// <summary>
        /// Gets or sets name of room.
        /// </summary>
        public string RoomName { get; set; }

        /// <summary>
        /// Gets or sets emails associated with room.
        /// </summary>
        public string RoomEmail { get; set; }

        /// <summary>
        /// Gets or sets meeting duration.
        /// </summary>
        public int Duration { get; set; }

        /// <summary>
        /// Gets or sets time zone.
        /// </summary>
        public string TimeZone { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether room is favorite.
        /// </summary>
        public bool IsFavourite { get; set; }

        /// <summary>
        /// Gets or sets time.
        /// </summary>
        public string Time { get; set; }
    }
}

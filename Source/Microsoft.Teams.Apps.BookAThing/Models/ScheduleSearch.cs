// <copyright file="ScheduleSearch.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Models
{
    /// <summary>
    /// Class containing properties required for search service.
    /// </summary>
    public class ScheduleSearch
    {
        /// <summary>
        /// Gets or sets search query.
        /// </summary>
        public string Query { get; set; }

        /// <summary>
        /// Gets or sets duration for which schedule to be fetched.
        /// </summary>
        public int Duration { get; set; }

        /// <summary>
        /// Gets or sets user local time zone.
        /// </summary>
        public string TimeZone { get; set; }

        /// <summary>
        /// Gets or sets time.
        /// </summary>
        public string Time { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether schedule required.
        /// </summary>
        public bool IsScheduleRequired { get; set; }
    }
}

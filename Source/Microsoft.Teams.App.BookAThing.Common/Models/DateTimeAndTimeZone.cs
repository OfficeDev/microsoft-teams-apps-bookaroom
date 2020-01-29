// <copyright file="DateTimeAndTimeZone.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Models
{
    using System;
    using Newtonsoft.Json;

    /// <summary>
    /// Date time with time zone.
    /// </summary>
    public class DateTimeAndTimeZone
    {
        /// <summary>
        /// Gets or sets date time.
        /// </summary>
        [JsonProperty("dateTime")]
        public DateTime DateTime { get; set; }

        /// <summary>
        /// Gets or sets time zone.
        /// </summary>
        [JsonProperty("timeZone")]
        public string TimeZone { get; set; }
    }
}

// <copyright file="Constants.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common
{
    using System;

    /// <summary>
    /// Constants class.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// Graph API base URL.
        /// </summary>
        public static readonly string GraphAPIBaseUrl = "https://graph.microsoft.com";

        /// <summary>
        /// Duration gap in minutes from now for which schedule for rooms will be fetched.
        /// </summary>
        public static readonly TimeSpan DurationGapFromNow = new TimeSpan(hours: 0, minutes: 5, seconds: 0);

        /// <summary>
        /// Default meeting duration in minutes.
        /// </summary>
        public static readonly TimeSpan DefaultMeetingDuration = new TimeSpan(hours: 0, minutes: 30, seconds: 0);
    }
}

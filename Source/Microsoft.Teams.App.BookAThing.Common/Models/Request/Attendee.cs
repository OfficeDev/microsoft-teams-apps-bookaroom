// <copyright file="Attendee.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Models.Request
{
    using Newtonsoft.Json;

    /// <summary>
    /// Attendee for meeting.
    /// </summary>
    public class Attendee
    {
        /// <summary>
        /// Gets or sets email address of attendee.
        /// </summary>
        [JsonProperty("emailAddress")]
        public EmailAddress EmailAddress { get; set; }

        /// <summary>
        /// Gets or sets attendee type.
        /// </summary>
        [JsonProperty("type")]
        public string Type { get; set; }
    }
}

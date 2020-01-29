// <copyright file="RoomScheduleResponse.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Models.Response
{
    using System.Collections.Generic;
    using System.Net;
    using Microsoft.Teams.App.BookAThing.Common.Models.Error;
    using Newtonsoft.Json;

    /// <summary>
    /// Schedule response.
    /// </summary>
    public class RoomScheduleResponse
    {
        /// <summary>
        /// Gets or sets list of schedule for a room.
        /// </summary>
        [JsonProperty("value")]
        public List<Schedule> Schedules { get; set; }

        /// <summary>
        /// Gets or sets Graph API response status code.
        /// </summary>
        public HttpStatusCode StatusCode { get; set; }

        /// <summary>
        /// Gets or sets Graph API error response.
        /// </summary>
        public ErrorResponse ErrorResponse { get; set; }
    }
}

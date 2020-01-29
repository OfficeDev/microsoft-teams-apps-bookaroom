// <copyright file="IMeetingHelper.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Helpers
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.Response;
    using Microsoft.Teams.Apps.BookAThing.Models;

    /// <summary>
    /// Exposes methods required for meeting creation.
    /// </summary>
    public interface IMeetingHelper
    {
        /// <summary>
        /// Create meeting.
        /// </summary>
        /// <param name="meeting">Object containing details required for meeting creation.</param>
        /// <param name="token">User Active Directory token.</param>
        /// <returns>CreateEventResponse object.</returns>
        Task<CreateEventResponse> CreateMeetingAsync(CreateEvent meeting, string token);

        /// <summary>
        /// Get rooms schedule.
        /// </summary>
        /// <param name="search">Object containing search query and time.</param>
        /// <param name="rooms">Room emails.</param>
        /// <param name="token">User Active Directory token.</param>
        /// <returns>List of schedule for rooms.</returns>
        Task<RoomScheduleResponse> GetRoomScheduleAsync(ScheduleSearch search, IList<string> rooms, string token);
    }
}
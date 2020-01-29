// <copyright file="IMeetingProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Providers
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.Request;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.Response;

    /// <summary>
    /// Provider which exposes methods required for meeting creation.
    /// </summary>
    public interface IMeetingProvider
    {
        /// <summary>
        /// Get schedule for rooms as per time selection.
        /// </summary>
        /// <param name="scheduleRequest">Schedule request object.</param>
        /// <param name="token">Active Directory access token.</param>
        /// <returns>Schedule response for given rooms.</returns>
        Task<RoomScheduleResponse> GetRoomsScheduleAsync(ScheduleRequest scheduleRequest, string token);

        /// <summary>
        /// Create new meeting for given room.
        /// </summary>
        /// <param name="eventRequest"><see cref="CreateEventRequest"/> object. </param>
        /// <param name="token">Active Directory access token.</param>
        /// <returns>Event response object.</returns>
        Task<CreateEventResponse> CreateMeetingAsync(CreateEventRequest eventRequest, string token);

        /// <summary>
        /// Cancel a meeting.
        /// </summary>
        /// <param name="meetingId">Unique meeting Id.</param>
        /// <param name="cancellationComment">Comment for meeting cancellation.</param>
        /// <param name="token">Active Directory access token.</param>
        /// <returns>Cancel meeting API response.</returns>
        Task<bool> CancelMeetingAsync(string meetingId, string cancellationComment, string token);
    }
}

// <copyright file="MeetingHelper.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.BookAThing.Common;
    using Microsoft.Teams.Apps.BookAThing.Common.Models;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.Request;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.Response;
    using Microsoft.Teams.Apps.BookAThing.Common.Providers;
    using Microsoft.Teams.Apps.BookAThing.Models;
    using Microsoft.Teams.Apps.BookAThing.Resources;

    /// <summary>
    /// Helper class which exposes methods required for meeting creation.
    /// </summary>
    public class MeetingHelper : IMeetingHelper
    {
        /// <summary>
        /// Provider for post and get API calls to Graph.
        /// </summary>
        private readonly IMeetingProvider meetingProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingHelper"/> class.
        /// </summary>
        /// <param name="meetingProvider">Provider for post and get API calls to Graph.</param>
        /// <param name="userConfigurationProvider">Provider for getting user specific configuration.</param>
        public MeetingHelper(IMeetingProvider meetingProvider)
        {
            this.meetingProvider = meetingProvider;
        }

        /// <summary>
        /// Get schedule for rooms.
        /// </summary>
        /// <param name="search">Object containing search query and time.</param>
        /// <param name="rooms">Room emails.</param>
        /// <param name="token">User Active Directory access token.</param>
        /// <returns>List of schedule for rooms.</returns>
        public async Task<RoomScheduleResponse> GetRoomScheduleAsync(ScheduleSearch search, IList<string> rooms, string token)
        {
            DateTime.TryParse(search.Time, out DateTime localTime);
            var startDateTime = localTime.AddMinutes(Constants.DurationGapFromNow.Minutes);
            var endDateTime = startDateTime.AddMinutes(Constants.DefaultMeetingDuration.Minutes);

            ScheduleRequest schedule = new ScheduleRequest
            {
                StartDateTime = new DateTimeAndTimeZone() { DateTime = startDateTime, TimeZone = search.TimeZone },
                EndDateTime = new DateTimeAndTimeZone() { DateTime = endDateTime, TimeZone = search.TimeZone },
                Schedules = new List<string>(),
            };

            schedule.Schedules.AddRange(rooms);
            return await this.meetingProvider.GetRoomsScheduleAsync(schedule, token).ConfigureAwait(false);
        }

        /// <summary>
        /// Create a meeting for selected time by user.
        /// </summary>
        /// <param name="meeting">Object containing details required for meeting creation.</param>
        /// <param name="token">User Active Directory access token.</param>
        /// <returns>Meeting response.</returns>
        public async Task<CreateEventResponse> CreateMeetingAsync(CreateEvent meeting, string token)
        {
            DateTime.TryParse(meeting.Time, out DateTime localTime);
            var startDateTime = localTime.AddMinutes(Constants.DurationGapFromNow.Minutes);
            var endDateTime = startDateTime.AddMinutes(meeting.Duration);

            ScheduleRequest scheduleRequest = new ScheduleRequest
            {
                StartDateTime = new DateTimeAndTimeZone() { DateTime = startDateTime, TimeZone = meeting.TimeZone },
                EndDateTime = new DateTimeAndTimeZone() { DateTime = endDateTime, TimeZone = meeting.TimeZone },
                Schedules = new List<string>(),
            };

            scheduleRequest.Schedules.Add(meeting.RoomEmail);
            var roomScheduleResponse = await this.meetingProvider.GetRoomsScheduleAsync(scheduleRequest, token).ConfigureAwait(false);

            if (roomScheduleResponse.ErrorResponse != null)
            {
                // Graph API returned error message.
                return new CreateEventResponse { StatusCode = roomScheduleResponse.StatusCode, ErrorResponse = roomScheduleResponse.ErrorResponse };
            }

            if (roomScheduleResponse.Schedules?.FirstOrDefault()?.ScheduleItems?.Count == 0)
            {
                CreateEventRequest request = new CreateEventRequest()
                {
                    Attendees = new List<Attendee>(),
                    Body = new Body { Content = Strings.MeetingBody, ContentType = "HTML" },
                    End = new DateTimeAndTimeZone { DateTime = endDateTime, TimeZone = meeting.TimeZone },
                    Start = new DateTimeAndTimeZone { DateTime = startDateTime, TimeZone = meeting.TimeZone },
                    Subject = Strings.MeetingBody,
                    Location = new Location { DisplayName = meeting.RoomName },
                };

                request.Attendees.Add(new Attendee { EmailAddress = new EmailAddress { Address = meeting.RoomEmail, Name = meeting.RoomName }, Type = "required" });
                return await this.meetingProvider.CreateMeetingAsync(request, token).ConfigureAwait(false);
            }
            else
            {
                return new CreateEventResponse
                {
                    ErrorResponse = new App.BookAThing.Common.Models.Error.ErrorResponse
                    {
                        Error = new App.BookAThing.Common.Models.Error.Error
                        {
                            StatusCode = "ScheduleExist",
                            ErrorMessage = "Schedule for room exist",
                        },
                    },
                };
            }
        }
    }
}

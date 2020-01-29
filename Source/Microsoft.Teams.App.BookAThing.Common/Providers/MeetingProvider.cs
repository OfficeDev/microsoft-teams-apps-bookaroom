// <copyright file="MeetingProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Teams.App.BookAThing.Common.Models.Error;
    using Microsoft.Teams.Apps.BookAThing.Common.Helpers;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.Request;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.Response;
    using Newtonsoft.Json;

    /// <summary>
    /// Exposes methods required for meeting creation.
    /// </summary>
    public class MeetingProvider : IMeetingProvider
    {
        /// <summary>
        /// Graph API events URL.
        /// </summary>
        private readonly string graphAPIEvents = "/v1.0/me/calendar/events";

        /// <summary>
        /// Graph API URL to get schedule for rooms.
        /// </summary>
        private readonly string graphAPIGetSchedule = "/v1.0/me/calendar/getSchedule";

        /// <summary>
        /// Graph API URL to cancel meeting. (Replace {id} with meeting Id).
        /// </summary>
        private readonly string graphAPICancelMeeting = "/beta/me/events/{id}/cancel";

        /// <summary>
        /// API helper service for making post and get calls to Graph.
        /// </summary>
        private readonly IGraphApiHelper apiHelper;

        /// <summary>
        /// Telemetry client to log event and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingProvider"/> class.
        /// </summary>
        /// <param name="apiHelper">Api helper service for making post and get calls to Graph.</param>
        /// <param name="telemetryClient">Telemetry client to log event and errors.</param>
        public MeetingProvider(IGraphApiHelper apiHelper, TelemetryClient telemetryClient)
        {
            this.apiHelper = apiHelper;
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Create new meeting for given room.
        /// </summary>
        /// <param name="eventRequest"><see cref="CreateEventRequest"/> object. </param>
        /// <param name="token">Active Directory access token.</param>
        /// <returns>Event response object.</returns>
        public async Task<CreateEventResponse> CreateMeetingAsync(CreateEventRequest eventRequest, string token)
        {
            var eventRequestPayload = JsonConvert.SerializeObject(eventRequest);
            Dictionary<string, string> header = new Dictionary<string, string>
            {
                { "Prefer", "outlook.timezone=\"" + eventRequest.End.TimeZone + "\"" },
            };

            var httpResponseMessage = await this.apiHelper.PostAsync(this.graphAPIEvents, token, eventRequestPayload, header).ConfigureAwait(false);
            var content = await httpResponseMessage.Content.ReadAsStringAsync();

            if (httpResponseMessage.IsSuccessStatusCode)
            {
                return JsonConvert.DeserializeObject<CreateEventResponse>(content);
            }

            var errorResponse = JsonConvert.DeserializeObject<ErrorResponse>(content);

            this.telemetryClient.TrackTrace($"Graph API failure- url: {this.graphAPIEvents}, response-code: {errorResponse.Error.StatusCode}, response-content: {errorResponse.Error.ErrorMessage}, request-id: {errorResponse.Error.InnerError.RequestId}", SeverityLevel.Warning);
            var failureResponse = new CreateEventResponse
            {
                StatusCode = httpResponseMessage.StatusCode,
                ErrorResponse = errorResponse,
            };

            return failureResponse;
        }

        /// <summary>
        /// Cancel a meeting.
        /// </summary>
        /// <param name="meetingId">Unique meeting id.</param>
        /// <param name="cancellationComment">Comment for cancellation of meeting.</param>
        /// <param name="token">Active Directory access token.</param>
        /// <returns>Cancel meeting API response.</returns>
        public async Task<bool> CancelMeetingAsync(string meetingId, string cancellationComment, string token)
        {
            var cancelMeetingJson = JsonConvert.SerializeObject(new { Comment = cancellationComment });
            var httpResponseMessage = await this.apiHelper.PostAsync(this.graphAPICancelMeeting.Replace("{id}", meetingId, StringComparison.OrdinalIgnoreCase), token, cancelMeetingJson).ConfigureAwait(false);
            var content = await httpResponseMessage.Content.ReadAsStringAsync();

            if (httpResponseMessage.IsSuccessStatusCode)
            {
                return true;
            }

            var errorResponse = JsonConvert.DeserializeObject<ErrorResponse>(content);
            this.telemetryClient.TrackTrace($"Graph API failure- url: {this.graphAPICancelMeeting.Replace("{id}", meetingId, StringComparison.OrdinalIgnoreCase)}, response-code: {errorResponse.Error.StatusCode}, response-content: {errorResponse.Error.ErrorMessage}, request-id: {errorResponse.Error.InnerError.RequestId}", SeverityLevel.Warning);
            return false;
        }

        /// <summary>
        /// Get schedule for rooms as per time selection.
        /// </summary>
        /// <param name="scheduleRequest">Schedule request object.</param>
        /// <param name="token">Active Directory access token.</param>
        /// <returns>Schedule of given rooms.</returns>
        public async Task<RoomScheduleResponse> GetRoomsScheduleAsync(ScheduleRequest scheduleRequest, string token)
        {
            var payload = JsonConvert.SerializeObject(scheduleRequest);
            Dictionary<string, string> header = new Dictionary<string, string>
            {
                { "Prefer", "outlook.timezone=\"" + scheduleRequest.StartDateTime.TimeZone + "\"" },
            };

            var httpResponseMessage = await this.apiHelper.PostAsync(this.graphAPIGetSchedule, token, payload, header).ConfigureAwait(false);
            var content = await httpResponseMessage.Content.ReadAsStringAsync();

            if (httpResponseMessage.IsSuccessStatusCode)
            {
                return JsonConvert.DeserializeObject<RoomScheduleResponse>(content);
            }

            var errorResponse = JsonConvert.DeserializeObject<ErrorResponse>(content);

            this.telemetryClient.TrackTrace($"Graph API failure- url: {this.graphAPIGetSchedule}, response-code: {errorResponse.Error.StatusCode}, response-content: {errorResponse.Error.ErrorMessage}, request-id: {errorResponse.Error.InnerError.RequestId}", SeverityLevel.Warning);
            var failureResponse = new RoomScheduleResponse
            {
                StatusCode = httpResponseMessage.StatusCode,
                ErrorResponse = errorResponse,
            };

            return failureResponse;
        }
    }
}

// <copyright file="MeetingApiController.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Threading.Tasks;

    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.App.BookAThing.Common.Models.Error;
    using Microsoft.Teams.Apps.BookAThing.Common;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.Response;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities;
    using Microsoft.Teams.Apps.BookAThing.Common.Providers;
    using Microsoft.Teams.Apps.BookAThing.Common.Providers.Storage;
    using Microsoft.Teams.Apps.BookAThing.Helpers;
    using Microsoft.Teams.Apps.BookAThing.Models;
    using Microsoft.Teams.Apps.BookAThing.Models.TableEntities;
    using Microsoft.Teams.Apps.BookAThing.Providers.Storage;
    using Microsoft.Teams.Apps.BookAThing.Resources;
    using TimeZoneConverter;

    /// <summary>
    /// Meeting API controller for handling API calls made from react js client app (used in task module).
    /// </summary>
    [ApiController]
    [Route("api/[controller]/[action]")]
    [Authorize]
    public class MeetingApiController : ControllerBase
    {
        /// <summary>
        /// Number of rooms to load in dropdown initially.
        /// </summary>
        private const int InitialRoomCount = 5;

        /// <summary>
        /// Search service for searching room/building as per user input.
        /// </summary>
        private readonly ISearchService searchService;

        /// <summary>
        /// Telemetry client to log event and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Generating and validating JWT token.
        /// </summary>
        private readonly ITokenHelper tokenHelper;

        /// <summary>
        /// Helper class which exposes methods required for meeting creation.
        /// </summary>
        private readonly IMeetingHelper meetingHelper;

        /// <summary>
        /// Storage provider to perform fetch, insert, update and delete operation on UserFavorites table.
        /// </summary>
        private readonly IFavoriteStorageProvider favoriteStorageProvider;

        /// <summary>
        /// Storage provider to perform fetch operation on RoomCollection table.
        /// </summary>
        private readonly IRoomCollectionStorageProvider roomCollectionStorageProvider;

        /// <summary>
        /// Storage provider to perform fetch operation on UserConfiguration table.
        /// </summary>
        private readonly IUserConfigurationStorageProvider userConfigurationStorageProvider;

        /// <summary>
        /// Provider to get user specific data from Graph API.
        /// </summary>
        private readonly IUserConfigurationProvider userConfigurationProvider;

        /// <summary>
        /// Unauthorized error message response in case of user sign in failure.
        /// </summary>
        private const string SignInErrorCode = "signinRequired";

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingApiController"/> class.
        /// </summary>
        /// <param name="favoriteStorageProvider">Storage provider to perform fetch, insert, update and delete operation on UserFavorites table.</param>
        /// <param name="roomCollectionStorageProvider">Storage provider to perform fetch operation on RoomCollection table.</param>
        /// <param name="searchService">Search service for searching room/building as per user input.</param>
        /// <param name="telemetryClient">Telemetry client to log event and errors.</param>
        /// <param name="tokenHelper">Generating and validating JWT token.</param>
        /// <param name="meetingHelper">Helper class which exposes methods required for meeting creation.</param>
        /// <param name="userConfigurationStorageProvider">Storage provider to perform fetch operation on UserConfiguration table.</param>
        /// <param name="userConfigurationProvider">Provider to get user specific data from Graph API.</param>
        public MeetingApiController(IFavoriteStorageProvider favoriteStorageProvider, IRoomCollectionStorageProvider roomCollectionStorageProvider, ISearchService searchService, TelemetryClient telemetryClient, ITokenHelper tokenHelper, IMeetingHelper meetingHelper, IUserConfigurationStorageProvider userConfigurationStorageProvider, IUserConfigurationProvider userConfigurationProvider)
        {
            this.favoriteStorageProvider = favoriteStorageProvider;
            this.roomCollectionStorageProvider = roomCollectionStorageProvider;
            this.searchService = searchService;
            this.telemetryClient = telemetryClient;
            this.tokenHelper = tokenHelper;
            this.meetingHelper = meetingHelper;
            this.userConfigurationStorageProvider = userConfigurationStorageProvider;
            this.userConfigurationProvider = userConfigurationProvider;
        }

        /// <summary>
        /// Get supported time zones for user from Graph API.
        /// </summary>
        /// <returns>Returns list of supported time zones.</returns>
        public async Task<IActionResult> GetSupportedTimeZonesAsync()
        {
            try
            {
                var claims = this.GetUserClaims();
                this.telemetryClient.TrackTrace($"User {claims.UserObjectIdentifer} submitted request to get supported time zones.");

                var token = await this.tokenHelper.GetUserTokenAsync(claims.FromId).ConfigureAwait(false);
                if (string.IsNullOrEmpty(token))
                {
                    this.telemetryClient.TrackTrace($"Azure Active Directory access token for user {claims.UserObjectIdentifer} is empty. Cannot get supported time zones.");
                    return this.StatusCode(
                        StatusCodes.Status401Unauthorized,
                        new Error
                        {
                            StatusCode = SignInErrorCode,
                            ErrorMessage = "Azure Active Directory access token for user is found empty.",
                        });
                }

                var supportedTimeZone = await this.userConfigurationProvider.GetSupportedTimeZoneAsync(token).ConfigureAwait(false);
                if (supportedTimeZone.ErrorResponse != null)
                {
                    // Graph API returned error message.
                    this.telemetryClient.TrackTrace($"Unable to fetch supported time zones for user {claims.UserObjectIdentifer}.");
                    return this.StatusCode((int)supportedTimeZone.StatusCode, supportedTimeZone.ErrorResponse.Error.ErrorMessage);
                }

                return this.Ok(supportedTimeZone.TimeZones?.OrderBy(timeZone => timeZone.DisplayName).Select(timeZone => timeZone.DisplayName));
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }

        /// <summary>
        /// Save user selected time zone.
        /// </summary>
        /// <param name="configuration">User configuration object.</param>
        /// <returns>Returns HTTP ok status code for successful operation.</returns>
        [HttpPost]
        public async Task<IActionResult> SaveTimeZoneAsync([FromBody]UserConfigurationEntity configuration)
        {
            try
            {
                var claims = this.GetUserClaims();
                this.telemetryClient.TrackTrace($"User {claims.UserObjectIdentifer} submitted settings. Time zone- {configuration.IanaTimezone}");
                configuration.UserAdObjectId = claims.UserObjectIdentifer;
                configuration.WindowsTimezone = TZConvert.IanaToWindows(configuration.IanaTimezone);

                var isAddOperationSuccess = await this.userConfigurationStorageProvider.AddAsync(configuration).ConfigureAwait(false);
                if (isAddOperationSuccess)
                {
                    return this.Ok("Configuration saved");
                }

                return this.StatusCode(StatusCodes.Status500InternalServerError, "Unable to save user configuration/time zone");
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }

        /// <summary>
        /// Retrieves user configuration settings.
        /// </summary>
        /// <returns>Returns user configuration settings.</returns>
        public async Task<IActionResult> GetUserTimeZoneAsync()
        {
            try
            {
                var claims = this.GetUserClaims();
                this.telemetryClient.TrackTrace($"User {claims.UserObjectIdentifer} requested time zone setting.");
                return this.Ok(await this.userConfigurationStorageProvider.GetAsync(claims.UserObjectIdentifer).ConfigureAwait(false));
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }

        /// <summary>
        ///  Get favorite rooms of user.
        /// </summary>
        /// <returns>Returns list of favourite rooms.</returns>
        public async Task<IActionResult> GetFavoriteRoomsAsync()
        {
            try
            {
                var claims = this.GetUserClaims();
                this.telemetryClient.TrackTrace($"User {claims.UserObjectIdentifer} opened task module and requested favorite rooms.");

                var token = await this.tokenHelper.GetUserTokenAsync(claims.FromId).ConfigureAwait(false);
                if (string.IsNullOrEmpty(token))
                {
                    this.telemetryClient.TrackTrace($"Azure Active Directory access token for user {claims.UserObjectIdentifer} is empty. Cannot get favorite rooms.");
                    return this.StatusCode(
                        StatusCodes.Status401Unauthorized,
                        new Error
                        {
                            StatusCode = SignInErrorCode,
                            ErrorMessage = "Azure Active Directory access token for user is found empty.",
                        });
                }

                var userFavoriteRooms = await this.favoriteStorageProvider.GetAsync(claims.UserObjectIdentifer).ConfigureAwait(false);
                return this.Ok(userFavoriteRooms);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }

        /// <summary>
        /// Get rooms/buildings using Azure search along with their schedule.
        /// </summary>
        /// <param name="search">Schedule object.</param>
        /// <returns>Returns list of rooms.</returns>
        public async Task<IActionResult> SearchRoomAsync([FromBody]ScheduleSearch search)
        {
            try
            {
                var claims = this.GetUserClaims();
                this.telemetryClient.TrackTrace($"Received search request for user {claims.UserObjectIdentifer}. Search query : {search?.Query}");

                var token = await this.tokenHelper.GetUserTokenAsync(claims.FromId).ConfigureAwait(false);
                if (string.IsNullOrEmpty(token))
                {
                    this.telemetryClient.TrackTrace($"Azure Active Directory access token for user {claims.UserObjectIdentifer} is empty. Cannot search rooms.");
                    return this.StatusCode(
                         StatusCodes.Status401Unauthorized,
                         new Error
                         {
                             StatusCode = SignInErrorCode,
                             ErrorMessage = "Azure Active Directory access token for user is found empty.",
                         });
                }

                var searchServiceResults = await this.searchService.SearchRoomsAsync(search.Query).ConfigureAwait(false);
                if (searchServiceResults == null)
                {
                    return this.StatusCode(StatusCodes.Status500InternalServerError);
                }

                if (search.IsScheduleRequired)
                {
                    var conversionResult = DateTime.TryParse(search.Time, out DateTime localTime);
                    var startUTCDateTime = localTime.AddMinutes(Constants.DurationGapFromNow.Minutes);
                    var startDateTime = TimeZoneInfo.ConvertTimeFromUtc(startUTCDateTime, TimeZoneInfo.FindSystemTimeZoneById(TZConvert.IanaToWindows(search.TimeZone)));
                    search.Time = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                    var rooms = searchServiceResults.Select(room => room.RowKey).ToList();
                    var scheduleResponse = await this.meetingHelper.GetRoomScheduleAsync(search, rooms, token).ConfigureAwait(false);

                    if (scheduleResponse.ErrorResponse != null)
                    {
                        return this.StatusCode((int)scheduleResponse.StatusCode, scheduleResponse.ErrorResponse.Error.ErrorMessage);
                    }

                    var searchedRooms = searchServiceResults.Select(searchResult => new SearchResult(searchResult)
                    {
                        Label = searchResult.RoomName,
                        Value = searchResult.RowKey,
                        Sublabel = searchResult.BuildingName,
                        Status = scheduleResponse.Schedules.Where(schedule => schedule.ScheduleId == searchResult.RowKey).FirstOrDefault()?.ScheduleItems?.Count > 0 ? Strings.Unavailable : Strings.Available,
                    }).ToList();

                    return this.Ok(searchedRooms);
                }
                else
                {
                    var searchedRooms = searchServiceResults.Select(searchResult => new SearchResult(searchResult)
                    {
                        Label = searchResult.RoomName,
                        Value = searchResult.RowKey,
                        Sublabel = searchResult.BuildingName,
                    }).ToList();

                    return this.Ok(searchedRooms);
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }

        /// <summary>
        /// Get top 'N' rooms from table storage where 'N' is room count.
        /// </summary>
        /// <param name="search">Schedule search object.</param>
        /// <returns>Returns list of rooms.</returns>
        public async Task<IActionResult> TopNRoomsAsync([FromBody]ScheduleSearch search)
        {
            try
            {
                var claims = this.GetUserClaims();
                var token = await this.tokenHelper.GetUserTokenAsync(claims.FromId).ConfigureAwait(false);
                if (string.IsNullOrEmpty(token))
                {
                    this.telemetryClient.TrackTrace($"Azure Active Directory access token for user {claims.UserObjectIdentifer} is empty. Cannot search rooms.");
                    return this.StatusCode(
                        (int)HttpStatusCode.Unauthorized,
                        new Error
                        {
                            StatusCode = SignInErrorCode,
                            ErrorMessage = "Azure Active Directory access token for user is found empty.",
                        });
                }

                var allRooms = await this.roomCollectionStorageProvider.GetNRoomsAsync(InitialRoomCount).ConfigureAwait(false);
                if (allRooms == null)
                {
                    return this.StatusCode(StatusCodes.Status500InternalServerError, "Unable to fetch rooms from storage");
                }

                if (search.IsScheduleRequired)
                {
                    var conversionResult = DateTime.TryParse(search.Time, out DateTime localTime);
                    var startUTCDateTime = localTime.AddMinutes(Constants.DurationGapFromNow.Minutes);
                    var startDateTime = TimeZoneInfo.ConvertTimeFromUtc(startUTCDateTime, TimeZoneInfo.FindSystemTimeZoneById(TZConvert.IanaToWindows(search.TimeZone)));
                    search.Time = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                    var rooms = allRooms.Select(room => room.RowKey).ToList();
                    var scheduleResponse = await this.meetingHelper.GetRoomScheduleAsync(search, rooms, token).ConfigureAwait(false);

                    if (scheduleResponse.ErrorResponse != null)
                    {
                        // Graph API returned error message.
                        return this.StatusCode((int)scheduleResponse.StatusCode, scheduleResponse.ErrorResponse.Error.ErrorMessage);
                    }

                    var searchResults = allRooms.Select(searchResult => new SearchResult(searchResult)
                    {
                        Label = searchResult.RoomName,
                        Value = searchResult.RowKey,
                        Sublabel = searchResult.BuildingName,
                        Status = scheduleResponse?.Schedules.Where(schedule => schedule.ScheduleId == searchResult.RowKey).FirstOrDefault()?.ScheduleItems?.Count > 0 ? Strings.Unavailable : Strings.Available,
                    }).ToList();

                    return this.Ok(searchResults);
                }
                else
                {
                    var searchResults = allRooms.Select(searchResult => new SearchResult(searchResult)
                    {
                        Label = searchResult.RoomName,
                        Value = searchResult.RowKey,
                        Sublabel = searchResult.BuildingName,
                    }).ToList();

                    return this.Ok(searchResults);
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }

        /// <summary>
        /// Add rooms to favorite in batch.
        /// </summary>
        /// <param name="rooms">List of rooms.</param>
        /// <returns>Returns HTTP ok status code for successful operation.</returns>
        [HttpPost]
        public async Task<IActionResult> SubmitFavoritesAsync([FromBody]IList<UserFavoriteRoomEntity> rooms)
        {
            try
            {
                var claims = this.GetUserClaims();
                this.telemetryClient.TrackTrace($"User {claims.UserObjectIdentifer} submitted favorite rooms");

                foreach (var room in rooms)
                {
                    room.UserAdObjectId = claims.UserObjectIdentifer;
                }

                var isDeleted = await this.favoriteStorageProvider.DeleteAllAsync(claims.UserObjectIdentifer).ConfigureAwait(false);
                if (!isDeleted)
                {
                    return this.StatusCode(StatusCodes.Status500InternalServerError, "Cannot delete favorite rooms");
                }

                if (rooms.Count > 0)
                {
                    var isInserted = await this.favoriteStorageProvider.AddBatchAsync(rooms).ConfigureAwait(false);
                    if (isInserted)
                    {
                        this.telemetryClient.TrackTrace($"Favorite rooms saved for user {claims.UserObjectIdentifer}");
                        return this.Ok("Favorite rooms saved");
                    }

                    return this.StatusCode(StatusCodes.Status500InternalServerError, "Cannot save favorite rooms");
                }

                return this.Ok();
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }

        /// <summary>
        /// Create meeting for selected room.
        /// </summary>
        /// <param name="meeting">Event object which will be sent to graph API.</param>
        /// <returns>Returns response receieved from Graph API containing Meeting Id, timing etc.</returns>
        [HttpPost]
        public async Task<ActionResult<CreateEventResponse>> CreateMeetingAsync([FromBody]CreateEvent meeting)
        {
            try
            {
                var claims = this.GetUserClaims();
                DateTime.TryParse(meeting.Time, out DateTime localTime);
                var startUTCDateTime = localTime.AddMinutes(Constants.DurationGapFromNow.Minutes);
                var endUTCDateTime = startUTCDateTime.AddMinutes(meeting.Duration);
                var startDateTime = TimeZoneInfo.ConvertTimeFromUtc(startUTCDateTime, TimeZoneInfo.FindSystemTimeZoneById(TZConvert.IanaToWindows(meeting.TimeZone)));
                meeting.Time = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");

                var token = await this.tokenHelper.GetUserTokenAsync(claims.FromId).ConfigureAwait(false);
                if (string.IsNullOrEmpty(token))
                {
                    this.telemetryClient.TrackTrace($"Azure Active Directory access token for user {claims.UserObjectIdentifer} is empty. Cannot create meeting.");
                    return this.StatusCode(
                        StatusCodes.Status401Unauthorized,
                        new Error
                        {
                            StatusCode = SignInErrorCode,
                            ErrorMessage = "Azure Active Directory access token for user is found empty.",
                        });
                }

                this.telemetryClient.TrackTrace($"User {claims.UserObjectIdentifer} initiated meeting creation for {meeting.RoomName}");
                var createEventResponse = await this.meetingHelper.CreateMeetingAsync(meeting, token).ConfigureAwait(false);

                if (createEventResponse.ErrorResponse != null)
                {
                    // Graph API returned error message.
                    this.telemetryClient.TrackTrace($"Meeting failed to create for user {claims.UserObjectIdentifer}. Room: {meeting.RoomName}, status-code: {createEventResponse.ErrorResponse.Error.StatusCode}, response-content: {createEventResponse.ErrorResponse.Error.ErrorMessage}");
                    return this.StatusCode((int)createEventResponse.StatusCode, createEventResponse.ErrorResponse.Error.ErrorMessage);
                }

                createEventResponse.Start.TimeZone = DateTime.SpecifyKind(startUTCDateTime, DateTimeKind.Utc).ToString("o");
                createEventResponse.End.TimeZone = DateTime.SpecifyKind(endUTCDateTime, DateTimeKind.Utc).ToString("o");
                this.telemetryClient.TrackTrace($"Meeting created for user {claims.UserObjectIdentifer}. Room - {meeting.RoomName}");
                return this.Ok(createEventResponse);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }

        /// <summary>
        /// Get claims of user.
        /// </summary>
        /// <returns>Claims.</returns>
        private JwtClaim GetUserClaims()
        {
            var claims = this.User.Claims;
            var jwtClaims = new JwtClaim
            {
                FromId = claims.Where(claim => claim.Type == "fromId").Select(claim => claim.Value).First(),
                ServiceUrl = claims.Where(claim => claim.Type == "serviceURL").Select(claim => claim.Value).First(),
                UserObjectIdentifer = claims.Where(claim => claim.Type == "userObjectIdentifer").Select(claim => claim.Value).First(),
            };

            return jwtClaims;
        }
    }
}
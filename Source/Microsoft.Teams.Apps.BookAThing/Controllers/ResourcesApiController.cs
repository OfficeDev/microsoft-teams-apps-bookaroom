// <copyright file="ResourcesApiController.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Controllers
{
    using System;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.BookAThing.Resources;

    /// <summary>
    /// Meeting API controller for handling API calls made from react js client app (used in task module).
    /// </summary>
    [ApiController]
    [Route("api/[controller]/[action]")]
    [Authorize]
    public class ResourcesApiController : ControllerBase
    {
        /// <summary>
        /// Telemetry client to log event and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="ResourcesApiController"/> class.
        /// </summary>
        /// <param name="telemetryClient">Telemetry client to log event and errors.</param>
        public ResourcesApiController(TelemetryClient telemetryClient)
        {
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Get resource strings for displaying in client app.
        /// </summary>
        /// <returns>Object containing required strings.</returns>
        public ActionResult GetResourceStrings()
        {
            try
            {
                var strings = new
                {
                    Strings.Timezone,
                    Strings.SelectTimezone,
                    Strings.LoadingMessage,
                    Strings.MeetingLength,
                    Strings.SearchRoom,
                    Strings.BookRoom,
                    Strings.InvalidTenant,
                    Strings.SearchRoomDropdownPlaceholder,
                    Strings.ExceptionResponse,
                    Strings.TimezoneNotSupported,
                    Strings.RoomUnavailable,
                    Strings.SelectDurationRoom,
                    Strings.Location,
                    Strings.AddButton,
                    Strings.DoneButton,
                    Strings.NoFavoriteRoomsTaskModule,
                    Strings.CantAddMoreRooms,
                    Strings.FavoriteRoomExist,
                    Strings.SelectRoomToAdd,
                    Strings.NoFavoritesDescriptionTaskModule,
                };
                return this.Ok(strings);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return this.StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }
    }
}
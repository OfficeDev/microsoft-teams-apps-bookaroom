// <copyright file="MainDialog.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Dialogs
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Dialogs;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.BookAThing.Cards;
    using Microsoft.Teams.Apps.BookAThing.Common;
    using Microsoft.Teams.Apps.BookAThing.Common.Models;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.Request;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.Response;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities;
    using Microsoft.Teams.Apps.BookAThing.Common.Providers;
    using Microsoft.Teams.Apps.BookAThing.Common.Providers.Storage;
    using Microsoft.Teams.Apps.BookAThing.Helpers;
    using Microsoft.Teams.Apps.BookAThing.Models;
    using Microsoft.Teams.Apps.BookAThing.Providers.Storage;
    using Microsoft.Teams.Apps.BookAThing.Resources;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Acts as root dialog for processing commands received from user.
    /// </summary>
    public class MainDialog : LogoutDialog
    {
        /// <summary>
        /// Helper which exposes methods required for meeting creation process.
        /// </summary>
        private readonly IMeetingProvider meetingProvider;

        /// <summary>
        /// Storage provider to perform fetch, insert and update operation on ActivityEntities table.
        /// </summary>
        private readonly IActivityStorageProvider activityStorageProvider;

        /// <summary>
        /// Storage provider to perform fetch, insert, update and delete operation on UserFavorites table.
        /// </summary>
        private readonly IFavoriteStorageProvider favoriteStorageProvider;

        /// <summary>
        /// Helper for generating and validating JWT token.
        /// </summary>
        private readonly ITokenHelper tokenHelper;

        /// <summary>
        /// Telemetry client for logging events and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Storage provider to perform fetch operation on UserConfiguration table.
        /// </summary>
        private readonly IUserConfigurationStorageProvider userConfigurationStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="MainDialog"/> class.
        /// </summary>
        /// <param name="configuration">Application configuration.</param>
        /// <param name="meetingProvider">Helper which exposes methods required for meeting creation process.</param>
        /// <param name="activityStorageProvider">Storage provider to perform fetch, insert and update operation on ActivityEntities table.</param>
        /// <param name="favoriteStorageProvider">Storage provider to perform fetch, insert, update and delete operation on UserFavorites table.</param>
        /// <param name="tokenHelper">Helper for generating and validating JWT token.</param>
        /// <param name="telemetryClient">Telemetry client for logging events and errors.</param>
        /// <param name="userConfigurationStorageProvider">Storage provider to perform fetch operation on UserConfiguration table.</param>
        public MainDialog(IConfiguration configuration, IMeetingProvider meetingProvider, IActivityStorageProvider activityStorageProvider, IFavoriteStorageProvider favoriteStorageProvider, ITokenHelper tokenHelper, TelemetryClient telemetryClient, IUserConfigurationStorageProvider userConfigurationStorageProvider)
            : base(nameof(MainDialog), configuration["ConnectionName"], telemetryClient)
        {
            this.tokenHelper = tokenHelper;
            this.telemetryClient = telemetryClient;
            this.activityStorageProvider = activityStorageProvider;
            this.favoriteStorageProvider = favoriteStorageProvider;
            this.userConfigurationStorageProvider = userConfigurationStorageProvider;
            this.meetingProvider = meetingProvider;
            this.AddDialog(new OAuthPrompt(
                 nameof(OAuthPrompt),
                 new OAuthPromptSettings
                 {
                     ConnectionName = this.ConnectionName,
                     Text = Strings.SignInRequired,
                     Title = Strings.SignIn,
                     Timeout = 120000,
                 }));
            this.AddDialog(
                new WaterfallDialog(
                    nameof(WaterfallDialog),
                    new WaterfallStep[] { this.PromptStepAsync, this.CommandStepAsync, this.ProcessStepAsync }));
            this.InitialDialogId = nameof(WaterfallDialog);
        }

        /// <summary>
        /// Prompts sign in card.
        /// </summary>
        /// <param name="stepContext">Context object passed in to a WaterfallStep.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task<DialogTurnResult> PromptStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            stepContext.Values["command"] = stepContext.Context.Activity.Text?.Trim();
            if (stepContext.Context.Activity.Text == null && stepContext.Context.Activity.Value != null && stepContext.Context.Activity.Type == "message")
            {
                stepContext.Values["command"] = JToken.Parse(stepContext.Context.Activity.Value.ToString()).SelectToken("text").ToString().Trim();
            }

            return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// To get access token, calling prompt again.
        /// </summary>
        /// <param name="stepContext">Context object passed in to a WaterfallStep.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task<DialogTurnResult> CommandStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var tokenResponse = (TokenResponse)stepContext.Result;
            if (!string.IsNullOrEmpty(tokenResponse?.Token))
            {
                if (stepContext.Values.ContainsKey("command"))
                {
                    stepContext.Context.Activity.Text = (string)stepContext.Values["command"] ?? string.Empty;
                }

                return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken).ConfigureAwait(false);
            }
            else
            {
                await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.CantLogIn), cancellationToken).ConfigureAwait(false);
                return await stepContext.EndDialogAsync().ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Process the command user typed.
        /// </summary>
        /// <param name="stepContext">Context object passed in to a WaterfallStep.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task<DialogTurnResult> ProcessStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            if (stepContext.Result != null)
            {
                var tokenResponse = stepContext.Result as TokenResponse;
                if (!string.IsNullOrEmpty(tokenResponse.Token))
                {
                    var command = (string)stepContext.Values["command"] ?? string.Empty;
                    switch (command.ToUpperInvariant())
                    {
                        case BotCommands.BookAMeeting:
                            await this.ShowRoomsAsync(stepContext, tokenResponse, false).ConfigureAwait(false);
                            break;

                        case BotCommands.RefreshList:
                            await this.ShowRoomsAsync(stepContext, tokenResponse, true).ConfigureAwait(false);
                            break;

                        case BotCommands.CreateMeeting:
                            await this.CreateEventAsync(stepContext, tokenResponse, cancellationToken).ConfigureAwait(false);
                            break;

                        case BotCommands.CancelMeeting:
                            await this.CancelMeetingAsync(stepContext, cancellationToken).ConfigureAwait(false);
                            break;

                        case BotCommands.AddFavorite:
                            await this.AddToFavoriteAsync(stepContext).ConfigureAwait(false);
                            break;

                        case BotCommands.ManageFavorites:
                            await this.ShowManageFavoritesCardAsync(stepContext).ConfigureAwait(false);
                            break;

                        case BotCommands.Help:
                            await this.ShowHelpCardAsync(stepContext).ConfigureAwait(false);
                            break;

                        case string message when message.Equals(BotCommands.Login, StringComparison.OrdinalIgnoreCase) || message.Equals(BotCommands.Logout, StringComparison.OrdinalIgnoreCase):
                            break;

                        default:
                            await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.CommandNotRecognized.Replace("{command}", command, StringComparison.CurrentCulture)), cancellationToken).ConfigureAwait(false);
                            break;
                    }
                }
            }
            else
            {
                await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.CantLogIn), cancellationToken).ConfigureAwait(false);
            }

            return await stepContext.EndDialogAsync().ConfigureAwait(false);
        }

        /// <summary>
        /// Show manage favorites card.
        /// </summary>
        /// <param name="stepContext">Context object passed in to a WaterfallStep.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task ShowManageFavoritesCardAsync(WaterfallStepContext stepContext)
        {
            var attachment = ManageFavoriteCard.GetManageFavoriteAttachment();
            await stepContext.Context.SendActivityAsync(MessageFactory.Attachment(attachment)).ConfigureAwait(false);
        }

        /// <summary>
        /// Send help card containing commands recognized by bot.
        /// </summary>
        /// <param name="stepContext">Context object passed in to a WaterfallStep.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task ShowHelpCardAsync(WaterfallStepContext stepContext)
        {
            var reply = stepContext.Context.Activity.CreateReply();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            reply.Attachments = HelpCard.GetHelpAttachments();
            await stepContext.Context.SendActivityAsync(reply).ConfigureAwait(false);
        }

        /// <summary>
        /// Display favorite rooms.
        /// </summary>
        /// <param name="stepContext">Context object passed in to a WaterfallStep.</param>
        /// <param name="tokenResponse">TokenResponse object containing user AAD token.</param>
        /// <param name="refresh">Boolean indicating whether call is for refreshing list.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task ShowRoomsAsync(WaterfallStepContext stepContext, TokenResponse tokenResponse, bool refresh)
        {
            var activity = stepContext.Context.Activity;
            var userFavorites = await this.favoriteStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);
            var startUTCTime = DateTime.UtcNow.AddMinutes(Constants.DurationGapFromNow.Minutes);
            var endUTCTime = startUTCTime.AddMinutes(Constants.DefaultMeetingDuration.Minutes);

            RoomScheduleResponse roomsScheduleResponse = new RoomScheduleResponse
            {
                Schedules = new List<Schedule>(),
            };

            var userConfiguration = await this.userConfigurationStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);
            if (userConfiguration == null)
            {
                this.telemetryClient.TrackTrace("User configuration is null in ShowRoomsAsync.");

                // Received empty user configuration but user has rooms in favorites table then show generic exception response and
                // return control back to caller. User must open task module to set time zone.
                if (userFavorites?.Count != 0)
                {
                    await stepContext.Context.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                    return;
                }
            }

            if (userFavorites?.Count > 0)
            {
                var startTime = TimeZoneInfo.ConvertTimeFromUtc(startUTCTime, TimeZoneInfo.FindSystemTimeZoneById(userConfiguration.WindowsTimezone));
                var endTime = startTime.AddMinutes(Constants.DefaultMeetingDuration.Minutes);
                roomsScheduleResponse = await this.GetRoomsScheduleAsync(startTime, endTime, userConfiguration.IanaTimezone, userFavorites, tokenResponse.Token).ConfigureAwait(false);

                if (roomsScheduleResponse.ErrorResponse != null)
                {
                    // Graph API returned error message.
                    await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.ExceptionResponse)).ConfigureAwait(false);
                    return;
                }

                foreach (var room in roomsScheduleResponse?.Schedules)
                {
                    var searchedRoom = userFavorites.Where(favoriteRoom => favoriteRoom.RowKey == room.ScheduleId).FirstOrDefault();
                    room.RoomName = searchedRoom.RoomName;
                    room.BuildingName = searchedRoom.BuildingName;
                }
            }

            // If user clicked refresh button.
            if (refresh)
            {
                var activityReferenceId = JObject.Parse(activity.Value.ToString()).SelectToken("activityReferenceId").ToString();
                var attachment = FavoriteRoomsListCard.GetFavoriteRoomsListAttachment(roomsScheduleResponse, startUTCTime, endUTCTime, userConfiguration?.WindowsTimezone, activityReferenceId);
                var updateCardActivity = new Activity(ActivityTypes.Message)
                {
                    Id = stepContext.Context.Activity.ReplyToId,
                    Conversation = stepContext.Context.Activity.Conversation,
                    Attachments = new List<Attachment> { attachment },
                };

                var replyActivity = await stepContext.Context.UpdateActivityAsync(updateCardActivity).ConfigureAwait(false);
                Models.TableEntities.ActivityEntity newActivity = new Models.TableEntities.ActivityEntity { ActivityId = replyActivity.Id, PartitionKey = activity.From.AadObjectId, RowKey = activityReferenceId };
                await this.activityStorageProvider.AddAsync(newActivity).ConfigureAwait(false);
            }
            else
            {
                var activityReferenceId = Guid.NewGuid().ToString();
                var attachment = FavoriteRoomsListCard.GetFavoriteRoomsListAttachment(roomsScheduleResponse, startUTCTime, endUTCTime, userConfiguration?.WindowsTimezone, activityReferenceId);
                var replyActivity = await stepContext.Context.SendActivityAsync(MessageFactory.Attachment(attachment)).ConfigureAwait(false);

                Models.TableEntities.ActivityEntity newActivity = new Models.TableEntities.ActivityEntity { ActivityId = replyActivity.Id, PartitionKey = activity.From.AadObjectId, RowKey = activityReferenceId };
                await this.activityStorageProvider.AddAsync(newActivity).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Add room to favorite from meeting success card.
        /// </summary>
        /// <param name="stepContext">Context object passed in to a WaterfallStep.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task AddToFavoriteAsync(WaterfallStepContext stepContext)
        {
            var activity = stepContext.Context.Activity;
            var postedValues = JsonConvert.DeserializeObject<MeetingViewModel>(activity.Value.ToString());
            var searchedFavoriteRooms = await this.favoriteStorageProvider.GetAsync(activity.From.AadObjectId, postedValues.RoomEmail).ConfigureAwait(false);

            if (searchedFavoriteRooms == null)
            {
                this.telemetryClient.TrackTrace("Searched favorite rooms is null in AddToFavoriteAsync.");
                await stepContext.Context.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
            }
            else if (searchedFavoriteRooms.Count > 0)
            {
                await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.FavoriteRoomExist)).ConfigureAwait(false);
            }
            else
            {
                UserFavoriteRoomEntity userFavorite = new UserFavoriteRoomEntity
                {
                    BuildingName = postedValues.BuildingName,
                    UserAdObjectId = activity.From.AadObjectId,
                    RoomEmail = postedValues.RoomEmail,
                    RoomName = postedValues.RoomName,
                    BuildingEmail = postedValues.BuildingEmail,
                };

                var favoriteRooms = await this.favoriteStorageProvider.AddAsync(userFavorite).ConfigureAwait(false);
                await stepContext.Context.SendActivityAsync(MessageFactory.Text(favoriteRooms?.Count > 0 ? Strings.RoomAddedAsFavorite : Strings.UnableToAddFavorite)).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Cancel a meeting.
        /// </summary>
        /// <param name="stepContext">Context object passed in to a WaterfallStep.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task CancelMeetingAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var activity = stepContext.Context.Activity;
            var tokenResponse = stepContext.Result as TokenResponse;
            var selectedMeetingValues = JsonConvert.DeserializeObject<MeetingViewModel>(activity.Value.ToString());
            var userConfiguration = await this.userConfigurationStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);
            if (userConfiguration == null)
            {
                this.telemetryClient.TrackTrace("User configuration is null in CancelMeetingAsync.");
                await stepContext.Context.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                return;
            }

            var startUTCDateTime = DateTime.Parse(selectedMeetingValues.StartDateTime, null, DateTimeStyles.RoundtripKind);
            var endUTCDateTime = DateTime.Parse(selectedMeetingValues.EndDateTime, null, DateTimeStyles.RoundtripKind);
            var endDateTime = TimeZoneInfo.ConvertTimeFromUtc(endUTCDateTime, TimeZoneInfo.FindSystemTimeZoneById(userConfiguration.WindowsTimezone));
            var localTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, TimeZoneInfo.FindSystemTimeZoneById(userConfiguration.WindowsTimezone));
            if (endDateTime < localTime)
            {
                await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.MeetingTimeElapsed)).ConfigureAwait(false);
            }
            else
            {
                var isCancelled = await this.meetingProvider.CancelMeetingAsync(selectedMeetingValues.MeetingId, Strings.CancellationComment, tokenResponse.Token).ConfigureAwait(false);

                if (isCancelled)
                {
                    var updateCardActivity = new Activity(ActivityTypes.Message)
                    {
                        Id = activity.ReplyToId,
                        Conversation = activity.Conversation,
                        Attachments = new List<Attachment> { CancellationCard.GetCancellationAttachment(selectedMeetingValues.RoomName, selectedMeetingValues.BuildingName, startUTCDateTime, endUTCDateTime, userConfiguration.WindowsTimezone) },
                    };

                    await stepContext.Context.UpdateActivityAsync(updateCardActivity, cancellationToken).ConfigureAwait(false);
                    await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.MeetingCancelled)).ConfigureAwait(false);
                }
                else
                {
                    await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.ExceptionResponse)).ConfigureAwait(false);
                }
            }
        }

        /// <summary>
        /// Create meeting as per room selection by user.
        /// </summary>
        /// <param name="stepContext">Context object passed in to a WaterfallStep.</param>
        /// <param name="tokenResponse">Token response object containing Active Directory access token.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task CreateEventAsync(WaterfallStepContext stepContext, TokenResponse tokenResponse, CancellationToken cancellationToken)
        {
            var activity = stepContext.Context.Activity;
            string available = "Available";

            var userConfiguration = await this.userConfigurationStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);
            if (userConfiguration == null)
            {
                this.telemetryClient.TrackTrace("User configuration is null in CreateEventAsync.");
                await stepContext.Context.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                return;
            }

            var selectedMeetingValues = JsonConvert.DeserializeObject<MeetingViewModel>(activity.Value.ToString());
            var startUTCDateTime = DateTime.Parse(selectedMeetingValues.StartDateTime, null, DateTimeStyles.RoundtripKind);
            var endUTCDateTime = DateTime.Parse(selectedMeetingValues.EndDateTime, null, DateTimeStyles.RoundtripKind);
            var currentDateTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, TimeZoneInfo.FindSystemTimeZoneById(userConfiguration.WindowsTimezone));
            var startDateTime = TimeZoneInfo.ConvertTimeFromUtc(startUTCDateTime, TimeZoneInfo.FindSystemTimeZoneById(userConfiguration.WindowsTimezone));
            var endDateTime = TimeZoneInfo.ConvertTimeFromUtc(endUTCDateTime, TimeZoneInfo.FindSystemTimeZoneById(userConfiguration.WindowsTimezone));

            if (currentDateTime.Subtract(startDateTime).TotalMinutes > Constants.DurationGapFromNow.Minutes)
            {
                await this.RefreshFavoriteCardAsync(stepContext).ConfigureAwait(false);
                await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.Expired), cancellationToken).ConfigureAwait(false);
            }
            else
            {
                if (!selectedMeetingValues.Status.Equals(available, StringComparison.OrdinalIgnoreCase))
                {
                    await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.RoomUnavailable), cancellationToken).ConfigureAwait(false);
                }
                else
                {
                    var roomsScheduleResponse = await this.GetRoomsScheduleAsync(
                        startDateTime,
                        endDateTime,
                        localTimeZone: userConfiguration.IanaTimezone,
                        rooms: new List<UserFavoriteRoomEntity>
                        {
                            new UserFavoriteRoomEntity { RowKey = selectedMeetingValues.RoomEmail },
                        },
                        tokenResponse.Token).ConfigureAwait(false);

                    if (roomsScheduleResponse.ErrorResponse != null)
                    {
                        // Graph API returned error message.
                        await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.ExceptionResponse)).ConfigureAwait(false);
                        return;
                    }

                    if (roomsScheduleResponse?.Schedules?.FirstOrDefault()?.ScheduleItems?.Count != 0)
                    {
                        await this.RefreshFavoriteCardAsync(stepContext).ConfigureAwait(false);
                        await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.RoomUnavailable), cancellationToken).ConfigureAwait(false);
                    }
                    else
                    {
                        CreateEventRequest request = new CreateEventRequest()
                        {
                            Attendees = new List<Attendee>(),
                            Body = new Body { Content = Strings.MeetingBody, ContentType = "HTML" },
                            End = new DateTimeAndTimeZone { DateTime = endDateTime, TimeZone = userConfiguration.IanaTimezone },
                            Start = new DateTimeAndTimeZone { DateTime = startDateTime, TimeZone = userConfiguration.IanaTimezone },
                            Subject = selectedMeetingValues.Subject,
                            Location = new Location { DisplayName = selectedMeetingValues.RoomName },
                        };

                        request.Attendees.Add(new Attendee { EmailAddress = new EmailAddress { Address = selectedMeetingValues.RoomEmail, Name = selectedMeetingValues.RoomName }, Type = "required" });
                        var meetingResponse = await this.meetingProvider.CreateMeetingAsync(request, tokenResponse.Token).ConfigureAwait(false);

                        if (meetingResponse.ErrorResponse != null)
                        {
                            // Graph API returned error message.
                            await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.ExceptionResponse), cancellationToken).ConfigureAwait(false);
                            return;
                        }

                        this.telemetryClient.TrackEvent("Meeting created", new Dictionary<string, string>() { { "User", activity.From.AadObjectId }, { "Room", selectedMeetingValues.RoomEmail } });
                        var updateCardActivity = new Activity(ActivityTypes.Message)
                        {
                            Id = stepContext.Context.Activity.ReplyToId,
                            Conversation = stepContext.Context.Activity.Conversation,
                            Attachments = new List<Attachment>
                                {
                                    SuccessCard.GetSuccessAttachment(
                                        new MeetingViewModel
                                        {
                                            MeetingId = meetingResponse.Id,
                                            RoomName = selectedMeetingValues.RoomName,
                                            BuildingName = selectedMeetingValues.BuildingName,
                                            WebLink = meetingResponse.WebLink,
                                            StartDateTime = selectedMeetingValues.StartDateTime,
                                            EndDateTime = selectedMeetingValues.EndDateTime,
                                            IsFavourite = true,
                                        },
                                        userConfiguration.WindowsTimezone),
                                },
                        };

                        await stepContext.Context.UpdateActivityAsync(updateCardActivity, cancellationToken).ConfigureAwait(false);
                        await stepContext.Context.SendActivityAsync(MessageFactory.Text(string.Format(CultureInfo.CurrentCulture, Strings.RoomBooked, selectedMeetingValues.RoomName)), cancellationToken).ConfigureAwait(false);
                    }
                }
            }
        }

        /// <summary>
        /// Update favorite list after task module close.
        /// </summary>
        /// <param name="stepContext">Context object passed in to a WaterfallStep.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task RefreshFavoriteCardAsync(WaterfallStepContext stepContext)
        {
            var activity = stepContext.Context.Activity;
            var userToken = await this.tokenHelper.GetUserTokenAsync(activity.From.Id).ConfigureAwait(false);
            var rooms = await this.favoriteStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);
            var activityReferenceId = Guid.NewGuid().ToString();
            var startUTCTime = DateTime.UtcNow.AddMinutes(Constants.DurationGapFromNow.Minutes);
            var endUTCTime = startUTCTime.AddMinutes(Constants.DefaultMeetingDuration.Minutes);
            var userConfiguration = await this.userConfigurationStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);

            if (userConfiguration == null)
            {
                this.telemetryClient.TrackTrace("User configuration is null in RefreshFavoriteCardAsync.");
                await stepContext.Context.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                return;
            }

            var startTime = TimeZoneInfo.ConvertTimeFromUtc(startUTCTime, TimeZoneInfo.FindSystemTimeZoneById(userConfiguration.WindowsTimezone));
            var endTime = startTime.AddMinutes(Constants.DefaultMeetingDuration.Minutes);

            if (rooms?.Count > 0)
            {
                var roomsScheduleResponse = await this.GetRoomsScheduleAsync(startTime, endTime, localTimeZone: userConfiguration.IanaTimezone, rooms, userToken).ConfigureAwait(false);
                if (roomsScheduleResponse.ErrorResponse != null)
                {
                    // Graph API returned error message.
                    await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.ExceptionResponse)).ConfigureAwait(false);
                    return;
                }

                foreach (var room in roomsScheduleResponse?.Schedules)
                {
                    var searchedRoom = rooms.Where(favoriteRoom => favoriteRoom.RowKey == room.ScheduleId).FirstOrDefault();
                    room.RoomName = searchedRoom.RoomName;
                    room.BuildingName = searchedRoom.BuildingName;
                }

                var attachment = FavoriteRoomsListCard.GetFavoriteRoomsListAttachment(roomsScheduleResponse, startUTCTime, endUTCTime, userConfiguration.WindowsTimezone, activityReferenceId);
                var updateCardActivity = new Activity(ActivityTypes.Message)
                {
                    Id = stepContext.Context.Activity.ReplyToId,
                    Conversation = stepContext.Context.Activity.Conversation,
                    Attachments = new List<Attachment> { attachment },
                };

                var replyActivity = await stepContext.Context.UpdateActivityAsync(updateCardActivity).ConfigureAwait(false);
                Models.TableEntities.ActivityEntity newActivity = new Models.TableEntities.ActivityEntity { ActivityId = replyActivity.Id, PartitionKey = activity.From.AadObjectId, RowKey = activityReferenceId };
                await this.activityStorageProvider.AddAsync(newActivity).ConfigureAwait(false);
            }
            else
            {
                RoomScheduleResponse scheduleResponse = new RoomScheduleResponse { Schedules = new List<Schedule>() };
                var attchment = FavoriteRoomsListCard.GetFavoriteRoomsListAttachment(scheduleResponse, startUTCTime, endUTCTime, userConfiguration.WindowsTimezone, activityReferenceId);

                var updateCardActivity = new Activity(ActivityTypes.Message)
                {
                    Id = stepContext.Context.Activity.ReplyToId,
                    Conversation = stepContext.Context.Activity.Conversation,
                    Attachments = new List<Attachment> { attchment },
                };

                var replyActivity = await stepContext.Context.UpdateActivityAsync(updateCardActivity).ConfigureAwait(false);
                Models.TableEntities.ActivityEntity newActivity = new Models.TableEntities.ActivityEntity { ActivityId = replyActivity.Id, PartitionKey = activity.From.AadObjectId, RowKey = activityReferenceId };
                await this.activityStorageProvider.AddAsync(newActivity).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Get schedule information of list of rooms.
        /// </summary>
        /// <param name="startTime">Start date time.</param>
        /// <param name="endTime">End date time.</param>
        /// <param name="localTimeZone">User local time zone.</param>
        /// <param name="rooms">List of rooms for which schedule needs to be fetched.</param>
        /// <param name="token">User Active Directory access token.</param>
        /// <returns>Room schedule response.</returns>
        private async Task<RoomScheduleResponse> GetRoomsScheduleAsync(DateTime startTime, DateTime endTime, string localTimeZone, IList<UserFavoriteRoomEntity> rooms, string token)
        {
            ScheduleRequest scheduleRequest = new ScheduleRequest
            {
                StartDateTime = new DateTimeAndTimeZone() { DateTime = startTime, TimeZone = localTimeZone },
                EndDateTime = new DateTimeAndTimeZone() { DateTime = endTime, TimeZone = localTimeZone },
                Schedules = new List<string>(),
            };

            scheduleRequest.Schedules.AddRange(rooms.Select(room => room.RoomEmail));
            return await this.meetingProvider.GetRoomsScheduleAsync(scheduleRequest, token).ConfigureAwait(false);
        }
    }
}

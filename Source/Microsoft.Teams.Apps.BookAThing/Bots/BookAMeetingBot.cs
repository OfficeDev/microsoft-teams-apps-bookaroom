// <copyright file="BookAMeetingBot.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Bots
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Dialogs;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
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
    /// Implements the core logic of the Book A Room bot.
    /// </summary>
    /// <typeparam name="T">Generic class.</typeparam>
    public class BookAMeetingBot<T> : TeamsActivityHandler
        where T : Dialog
    {
        /// <summary>
        /// Reads and writes conversation state for your bot to storage.
        /// </summary>
        private readonly BotState conversationState;

        /// <summary>
        /// Dialog to be invoked.
        /// </summary>
        private readonly Dialog dialog;

        /// <summary>
        /// Stores user specific data.
        /// </summary>
        private readonly BotState userState;

        /// <summary>
        /// Application base URL.
        /// </summary>
        private readonly string appBaseUri;

        /// <summary>
        /// Instrumentation key for application insights logging.
        /// </summary>
        private readonly string instrumentationKey;

        /// <summary>
        /// Valid tenant id for which bot will operate.
        /// </summary>
        private readonly string tenantId;

        /// <summary>
        /// Generating and validating JWT token.
        /// </summary>
        private readonly ITokenHelper tokenHelper;

        /// <summary>
        /// Storage provider to perform insert, update and delete operation on ActivityEntities table.
        /// </summary>
        private readonly IActivityStorageProvider activityStorageProvider;

        /// <summary>
        /// Storage provider to perform insert, update and delete operation on UserFavorites table.
        /// </summary>
        private readonly IFavoriteStorageProvider favoriteStorageProvider;

        /// <summary>
        /// Provider for exposing methods required to perform meeting creation.
        /// </summary>
        private readonly IMeetingProvider meetingProvider;

        /// <summary>
        /// Telemetry client for logging events and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Storage provider to perform insert and update operation on UserConfiguration table.
        /// </summary>
        private readonly IUserConfigurationStorageProvider userConfigurationStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="BookAMeetingBot{T}"/> class.
        /// </summary>
        /// <param name="conversationState">Reads and writes conversation state for your bot to storage.</param>
        /// <param name="userState">Reads and writes user specific data to storage.</param>
        /// <param name="dialog">Dialog to be invoked.</param>
        /// <param name="tokenHelper">Generating and validating JWT token.</param>
        /// <param name="activityStorageProvider">Storage provider to perform insert, update and delete operation on ActivityEntities table.</param>
        /// <param name="favoriteStorageProvider">Storage provider to perform insert, update and delete operation on UserFavorites table.</param>
        /// <param name="meetingProvider">Provider for exposing methods required to perform meeting creation.</param>
        /// <param name="telemetryClient">Telemetry client for logging events and errors.</param>
        /// <param name="userConfigurationStorageProvider">Storage provider to perform insert and update operation on UserConfiguration table.</param>
        /// <param name="appBaseUri">Application base URL.</param>
        /// <param name="instrumentationKey">Instrumentation key for application insights logging.</param>
        /// <param name="tenantId">Valid tenant id for which bot will operate.</param>
        public BookAMeetingBot(ConversationState conversationState, UserState userState, T dialog, ITokenHelper tokenHelper, IActivityStorageProvider activityStorageProvider, IFavoriteStorageProvider favoriteStorageProvider, IMeetingProvider meetingProvider, TelemetryClient telemetryClient, IUserConfigurationStorageProvider userConfigurationStorageProvider, string appBaseUri, string instrumentationKey, string tenantId)
        {
            this.conversationState = conversationState;
            this.userState = userState;
            this.dialog = dialog;
            this.tokenHelper = tokenHelper;
            this.activityStorageProvider = activityStorageProvider;
            this.favoriteStorageProvider = favoriteStorageProvider;
            this.meetingProvider = meetingProvider;
            this.telemetryClient = telemetryClient;
            this.userConfigurationStorageProvider = userConfigurationStorageProvider;
            this.appBaseUri = appBaseUri;
            this.instrumentationKey = instrumentationKey;
            this.tenantId = tenantId;
        }

        /// <summary>
        /// Handles an incoming activity.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public override async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default)
        {
            var activity = turnContext.Activity;
            if (!this.IsActivityFromExpectedTenant(turnContext))
            {
                this.telemetryClient.TrackTrace($"Unexpected tenant Id {activity.Conversation.TenantId}", SeverityLevel.Warning);
                await turnContext.SendActivityAsync(MessageFactory.Text(Strings.InvalidTenant)).ConfigureAwait(false);
            }
            else
            {
                this.telemetryClient.TrackTrace($"Activity received = Activity Id: {activity.Id}, Activity type: {activity.Type}, Activity text: {activity?.Text}, From Id: {activity.From?.Id}, User object Id: {activity.From?.AadObjectId}", SeverityLevel.Information);
                var locale = activity.Entities?.Where(entity => entity.Type == "clientInfo").First().Properties["locale"].ToString();

                // Get the current culture info to use in resource files
                if (locale != null)
                {
                    CultureInfo.CurrentUICulture = CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(locale);
                }

                await base.OnTurnAsync(turnContext, cancellationToken).ConfigureAwait(false);

                // Save any state changes that might have occured during the turn.
                await this.conversationState.SaveChangesAsync(turnContext, false, cancellationToken).ConfigureAwait(false);
                await this.userState.SaveChangesAsync(turnContext, false, cancellationToken).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Invoked when members other than this bot (like a user) are added to the conversation.
        /// </summary>
        /// <param name="membersAdded">List of members added.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var activity = turnContext?.Activity;
            this.telemetryClient.TrackTrace($"conversationType: {activity.Conversation.ConversationType}, membersAdded: {activity.MembersAdded?.Count}, membersRemoved: {activity.MembersRemoved?.Count}");

            if (activity.MembersAdded?.Where(member => member.Id != activity.Recipient.Id).FirstOrDefault() != null)
            {
                this.telemetryClient.TrackEvent("Bot installed", new Dictionary<string, string>() { { "User", activity.From.AadObjectId } });
                var userStateAccessors = this.userState.CreateProperty<UserData>(nameof(UserData));
                var userdata = await userStateAccessors.GetAsync(turnContext, () => new UserData()).ConfigureAwait(false);

                if (userdata?.IsWelcomeCardSent == null || userdata?.IsWelcomeCardSent == false)
                {
                    userdata.IsWelcomeCardSent = true;
                    await userStateAccessors.SetAsync(turnContext, userdata).ConfigureAwait(false);
                    var welcomeCardImageUrl = new Uri(baseUri: new Uri(this.appBaseUri), relativeUri: "/images/welcome.jpg");
                    await turnContext.SendActivityAsync(activity: MessageFactory.Attachment(WelcomeCard.GetWelcomeCardAttachment(welcomeCardImageUrl)), cancellationToken).ConfigureAwait(false);
                }
            }
        }

        /// <summary>
        /// Invoked when members other than this bot (like a user) are removed from the conversation.
        /// </summary>
        /// <param name="membersRemoved">List of members removed.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnMembersRemovedAsync(IList<ChannelAccount> membersRemoved, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var activity = turnContext?.Activity;
            this.telemetryClient.TrackTrace($"conversationType: {activity.Conversation?.ConversationType}, membersAdded: {activity.MembersAdded?.Count}, membersRemoved: {activity.MembersRemoved?.Count}");

            if (activity.MembersAdded?.Where(member => member.Id != activity.Recipient.Id).FirstOrDefault() != null)
            {
                this.telemetryClient.TrackEvent("Bot uninstalled", new Dictionary<string, string>() { { "User", activity.From.AadObjectId } });
                var userStateAccessors = this.userState.CreateProperty<UserData>(nameof(UserData));
                var userdata = await userStateAccessors.GetAsync(turnContext, () => new UserData()).ConfigureAwait(false);
                userdata.IsWelcomeCardSent = false;
                await userStateAccessors.SetAsync(turnContext, userdata).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Invoked when a signin or verify activity is received.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnTeamsSigninVerifyStateAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            if (turnContext.Activity.Name.Equals("signin/verifyState", StringComparison.OrdinalIgnoreCase))
            {
                await turnContext.SendActivityAsync(MessageFactory.Text(Strings.LoggedInSuccess), cancellationToken).ConfigureAwait(false);
            }

            await this.dialog.RunAsync(turnContext, this.conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Invoked when task module fetch event is received from the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            var activity = turnContext.Activity;
            if (taskModuleRequest.Data == null)
            {
                this.telemetryClient.TrackTrace("Request data obtained on task module fetch action is null.");
                await turnContext.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                return default;
            }

            var userToken = await this.tokenHelper.GetUserTokenAsync(activity.From.Id);
            if (string.IsNullOrEmpty(userToken))
            {
                // No token found for user. Trying to open task module after sign out.
                this.telemetryClient.TrackTrace("User token is null in OnTeamsTaskModuleFetchAsync.");
                await turnContext.SendActivityAsync(Strings.SignInErrorMessage).ConfigureAwait(false);
                return default;
            }

            var postedValues = JsonConvert.DeserializeObject<MeetingViewModel>(JObject.Parse(taskModuleRequest.Data.ToString()).SelectToken("data").ToString());
            var command = postedValues.Text;
            var token = this.tokenHelper.GenerateAPIAuthToken(activity.From.AadObjectId, activity.ServiceUrl, activity.From.Id, jwtExpiryMinutes: 60);
            string activityReferenceId = string.Empty;

            switch (command.ToUpperInvariant())
            {
                // Show task module to manage favorite rooms which is invoked from 'Favorite rooms' list card.
                case BotCommands.ShowFavoriteTaskModule:
                    activityReferenceId = postedValues.ActivityReferenceId;
                    return this.GetTaskModuleResponse(string.Format(CultureInfo.InvariantCulture, "{0}/Meeting/AddFavourite?telemetry={1}&token={2}&replyTo={3}", this.appBaseUri, this.instrumentationKey, token, activityReferenceId), Strings.AddFavTaskModuleSubtitle);

                // Show task module to manage favorite rooms which is invoked from 'Manage favorites' card.
                case BotCommands.ManageFavorites:
                    activityReferenceId = string.Empty;
                    return this.GetTaskModuleResponse(string.Format(CultureInfo.InvariantCulture, "{0}/Meeting/AddFavourite?telemetry={1}&token={2}&replyTo={3}", this.appBaseUri, this.instrumentationKey, token, activityReferenceId), Strings.AddFavTaskModuleSubtitle);

                // Show task module to book room which is not added in favorites.
                case BotCommands.ShowOtherRoomsTaskModule:
                    activityReferenceId = postedValues.ActivityReferenceId;
                    return this.GetTaskModuleResponse(string.Format(CultureInfo.InvariantCulture, "{0}/Meeting/OtherRoom?telemetry={1}&token={2}&replyTo={3}", this.appBaseUri, this.instrumentationKey, token, activityReferenceId), Strings.AnotherRoomTaskModuleSubtitle);

                default:
                    var reply = MessageFactory.Text(Strings.CommandNotRecognized.Replace("{command}", command, StringComparison.OrdinalIgnoreCase));
                    await turnContext.SendActivityAsync(reply).ConfigureAwait(false);
                    return default;
            }
        }

        /// <summary>
        /// Invoked when task module submit event is received from the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            var activity = turnContext.Activity;
            if (taskModuleRequest.Data == null)
            {
                this.telemetryClient.TrackTrace("Request data obtained on task module submit action is null.");
                await turnContext.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                return default;
            }

            var userConfiguration = await this.userConfigurationStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);
            if (userConfiguration == null)
            {
                this.telemetryClient.TrackTrace("User configuration is null in task module submit action.");
                await turnContext.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                return default;
            }

            var valuesFromTaskModule = JsonConvert.DeserializeObject<MeetingViewModel>(taskModuleRequest.Data.ToString());
            var message = valuesFromTaskModule.Text;
            var replyToId = valuesFromTaskModule.ReplyTo;

            if (message.Equals(BotCommands.MeetingFromTaskModule, StringComparison.OrdinalIgnoreCase))
            {
                var attachment = SuccessCard.GetSuccessAttachment(valuesFromTaskModule, userConfiguration.WindowsTimezone);
                var activityFromStorage = await this.activityStorageProvider.GetAsync(activity.From.AadObjectId, replyToId).ConfigureAwait(false);

                if (!string.IsNullOrEmpty(replyToId))
                {
                    var updateCardActivity = new Activity(ActivityTypes.Message)
                    {
                        Id = activityFromStorage.ActivityId,
                        Conversation = activity.Conversation,
                        Attachments = new List<Attachment> { attachment },
                    };
                    await turnContext.UpdateActivityAsync(updateCardActivity).ConfigureAwait(false);
                }

                await turnContext.SendActivityAsync(MessageFactory.Text(string.Format(CultureInfo.CurrentCulture, Strings.RoomBooked, valuesFromTaskModule.RoomName)), cancellationToken).ConfigureAwait(false);
            }
            else
            {
                if (!string.IsNullOrEmpty(replyToId))
                {
                    await this.UpdateFavouriteCardAsync(turnContext, replyToId).ConfigureAwait(false);
                }
            }

            return null;
        }

        /// <summary>
        /// Invoked when a message activity is received from the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var command = turnContext.Activity.Text;
            await SendTypingIndicatorAsync(turnContext).ConfigureAwait(false);

            if (turnContext.Activity.Text == null && turnContext.Activity.Value != null && turnContext.Activity.Type == ActivityTypes.Message)
            {
                command = JToken.Parse(turnContext.Activity.Value.ToString()).SelectToken("text").ToString();
            }

            switch (command.ToUpperInvariant())
            {
                case BotCommands.Help:
                    await ShowHelpCardAsync(turnContext).ConfigureAwait(false);
                    break;

                default:
                    await this.dialog.RunAsync(turnContext, this.conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken).ConfigureAwait(false);
                    break;
            }
        }

        /// <summary>
        /// Send help card containing commands recognized by bot.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private static async Task ShowHelpCardAsync(ITurnContext<IMessageActivity> turnContext)
        {
            var activity = (Activity)turnContext.Activity;
            var reply = activity.CreateReply();
            reply.Attachments = HelpCard.GetHelpAttachments();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            await turnContext.SendActivityAsync(reply).ConfigureAwait(false);
        }

        /// <summary>
        /// Send typing indicator to the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private static Task SendTypingIndicatorAsync(ITurnContext turnContext)
        {
            var typingActivity = turnContext.Activity.CreateReply();
            typingActivity.Type = ActivityTypes.Typing;
            return turnContext.SendActivityAsync(typingActivity);
        }

        /// <summary>
        /// Update favorite list after task module close.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="activityReferenceId">Unique GUID related to activity Id from ActivityEntities table.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task UpdateFavouriteCardAsync(ITurnContext<IInvokeActivity> turnContext, string activityReferenceId)
        {
            var activity = turnContext.Activity;
            var userAADToken = await this.tokenHelper.GetUserTokenAsync(activity.From.Id).ConfigureAwait(false);
            var userConfiguration = await this.userConfigurationStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);
            if (userConfiguration == null)
            {
                this.telemetryClient.TrackTrace("User configuration is null in UpdateFavouriteCardAsync.");
                await turnContext.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                return;
            }

            var rooms = await this.favoriteStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);
            var startUTCTime = DateTime.UtcNow.AddMinutes(Constants.DurationGapFromNow.Minutes);
            var startTime = TimeZoneInfo.ConvertTimeFromUtc(startUTCTime, TimeZoneInfo.FindSystemTimeZoneById(userConfiguration.WindowsTimezone));
            var endTime = startTime.AddMinutes(Constants.DefaultMeetingDuration.Minutes);

            if (rooms?.Count > 0)
            {
                ScheduleRequest request = new ScheduleRequest
                {
                    StartDateTime = new DateTimeAndTimeZone() { DateTime = startTime, TimeZone = userConfiguration.IanaTimezone },
                    EndDateTime = new DateTimeAndTimeZone() { DateTime = endTime, TimeZone = userConfiguration.IanaTimezone },
                    Schedules = new List<string>(),
                };

                request.Schedules.AddRange(rooms.Select(room => room.RoomEmail));
                var roomsScheduleResponse = await this.meetingProvider.GetRoomsScheduleAsync(request, userAADToken).ConfigureAwait(false);
                if (roomsScheduleResponse.ErrorResponse == null)
                {
                    await this.SendAndUpdateCardAsync(turnContext, rooms, roomsScheduleResponse, activityReferenceId).ConfigureAwait(false);
                }
                else
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text(Strings.ExceptionResponse)).ConfigureAwait(false);
                }
            }
            else
            {
                RoomScheduleResponse scheduleResponse = new RoomScheduleResponse { Schedules = new List<Schedule>() };
                await this.SendAndUpdateCardAsync(turnContext, rooms, scheduleResponse, activityReferenceId).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Update list card containing user favorites with latest room availability.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="rooms">List of favorite rooms for a user.</param>
        /// <param name="scheduleResponse">Schedule received for favorite rooms of user.</param>
        /// <param name="activityReferenceId">Unique GUID related to activity Id from ActivityEntities table.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task SendAndUpdateCardAsync(ITurnContext<IInvokeActivity> turnContext, IList<UserFavoriteRoomEntity> rooms, RoomScheduleResponse scheduleResponse, string activityReferenceId)
        {
            var activity = turnContext.Activity;
            var startUTCTime = DateTime.UtcNow.AddMinutes(Constants.DurationGapFromNow.Minutes);
            var endUTCTime = startUTCTime.AddMinutes(Constants.DefaultMeetingDuration.Minutes);
            var userConfiguration = await this.userConfigurationStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);
            if (userConfiguration == null)
            {
                this.telemetryClient.TrackTrace("User configuration is null in SendAndUpdateCardAsync.");
                await turnContext.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                return;
            }

            var startTime = TimeZoneInfo.ConvertTimeFromUtc(startUTCTime, TimeZoneInfo.FindSystemTimeZoneById(userConfiguration.WindowsTimezone));

            foreach (var room in scheduleResponse.Schedules)
            {
                var searchedRoom = rooms.Where(favoriteRoom => favoriteRoom.RowKey == room.ScheduleId).FirstOrDefault();
                room.RoomName = searchedRoom?.RoomName;
                room.BuildingName = searchedRoom?.BuildingName;
            }

            var activityFromStorage = await this.activityStorageProvider.GetAsync(turnContext.Activity.From.AadObjectId, activityReferenceId).ConfigureAwait(false);
            if (activityFromStorage != null)
            {
                var attachment = FavoriteRoomsListCard.GetFavoriteRoomsListAttachment(scheduleResponse, startUTCTime, endUTCTime, userConfiguration.WindowsTimezone, activityReferenceId);
                var updateCardActivity = new Activity(ActivityTypes.Message)
                {
                    Id = activityFromStorage.ActivityId,
                    Conversation = turnContext.Activity.Conversation,
                    Attachments = new List<Attachment> { attachment },
                };

                var activityResponse = await turnContext.UpdateActivityAsync(updateCardActivity).ConfigureAwait(false);
                Models.TableEntities.ActivityEntity newActivity = new Models.TableEntities.ActivityEntity { ActivityId = activityResponse.Id, PartitionKey = turnContext.Activity.From.AadObjectId, RowKey = activityReferenceId };
                await this.activityStorageProvider.AddAsync(newActivity).ConfigureAwait(false);
            }
            else
            {
                await turnContext.SendActivityAsync(MessageFactory.Text(Strings.FavoriteRoomsModified)).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Verify if the tenant Id in the message is the same tenant Id used when application was configured.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>Boolean indicating whether tenant is valid.</returns>
        private bool IsActivityFromExpectedTenant(ITurnContext turnContext)
        {
            return turnContext.Activity.Conversation.TenantId.Equals(this.tenantId, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Get task module response object.
        /// </summary>
        /// <param name="url">Task module URL.</param>
        /// <param name="title">Title for task module.</param>
        /// <returns>TaskModuleResponse object.</returns>
        private TaskModuleResponse GetTaskModuleResponse(string url, string title)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Type = "continue",
                    Value = new TaskModuleTaskInfo()
                    {
                        Url = url,
                        Height = 460,
                        Width = 600,
                        Title = title,
                    },
                },
            };
        }
    }
}
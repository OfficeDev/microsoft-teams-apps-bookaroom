// <copyright file="FavoriteRoomsListCard.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.Response;
    using Microsoft.Teams.Apps.BookAThing.Models;
    using Microsoft.Teams.Apps.BookAThing.Resources;
    using Newtonsoft.Json;

    /// <summary>
    /// Class having method to get favorites list.
    /// </summary>
    public static class FavoriteRoomsListCard
    {
        /// <summary>
        /// Color code for unavailable status.
        /// </summary>
        private const string RedColorCode = "#CC4A31";

        /// <summary>
        /// Color code for available status.
        /// </summary>
        private const string GreenColorCode = "#92C353";

        /// <summary>
        /// Get list card attachment having list of favorite rooms along with buttons to manage favorites, book meeting for other rooms and refresh list.
        /// </summary>
        /// <param name="rooms">Rooms schedule response object.</param>
        /// <param name="startUTCDateTime">Start date time of meeting.</param>
        /// <param name="endUTCDateTime">End date time of meeting.</param>
        /// <param name="timeZone">User time zone.</param>
        /// <param name="activityReferenceId">Unique GUID related to activity Id from ActivityEntities table.</param>
        /// <returns>List card attachment having favorite rooms of user.</returns>
        public static Attachment GetFavoriteRoomsListAttachment(RoomScheduleResponse rooms, DateTime startUTCDateTime, DateTime endUTCDateTime, string timeZone, string activityReferenceId = null)
        {
            ListCard card = new ListCard
            {
                Title = Strings.RoomAvailability,
                Items = new List<ListItem>(),
                Buttons = new List<CardAction>(),
            };

            // For first run, user configuration will be null.
            if (timeZone != null)
            {
                var startTime = TimeZoneInfo.ConvertTimeFromUtc(startUTCDateTime, TimeZoneInfo.FindSystemTimeZoneById(timeZone));
                var endTime = TimeZoneInfo.ConvertTimeFromUtc(endUTCDateTime, TimeZoneInfo.FindSystemTimeZoneById(timeZone));
                card.Title = string.Format(CultureInfo.CurrentCulture, "{0} | {1} - {2}", Strings.RoomAvailability, startTime.ToString("t", CultureInfo.CurrentCulture), endTime.ToString("t", CultureInfo.CurrentCulture));
            }

            ListItem room = new ListItem();
            Meeting meeting;

            if (rooms.Schedules.Count > 0)
            {
                foreach (var item in rooms.Schedules)
                {
                    var availability = item.ScheduleItems.Count > 0 ? Strings.Unavailable : Strings.Available;
                    var availabilityColor = item.ScheduleItems.Count > 0 ? RedColorCode : GreenColorCode;
                    var subtitle = string.Format(CultureInfo.CurrentCulture, "{0}&nbsp;|&nbsp;<b><font color='{1}'>{2}</font></b>", item.BuildingName, availabilityColor, availability);

                    meeting = new Meeting
                    {
                        EndDateTime = DateTime.SpecifyKind(endUTCDateTime, DateTimeKind.Utc).ToString("o"),
                        RoomEmail = item.ScheduleId,
                        RoomName = item.RoomName,
                        StartDateTime = DateTime.SpecifyKind(startUTCDateTime, DateTimeKind.Utc).ToString("o"),
                        BuildingName = item.BuildingName,
                        Status = availability,
                        Text = BotCommands.CreateMeeting,
                    };

                    card.Items.Add(new ListItem
                    {
                        Id = item.ScheduleId,
                        Title = item.RoomName,
                        Subtitle = subtitle,
                        Type = "person",
                        Tap = new CardAction { Type = ActionTypes.MessageBack, DisplayText = string.Empty, Title = BotCommands.CreateMeeting, Value = JsonConvert.SerializeObject(meeting) },
                    });
                }
            }
            else
            {
                room = new ListItem
                {
                    Title = Strings.NoFavoriteRooms,
                    Type = "section",
                };
            }

            card.Items.Add(room);
            CardAction addFavoriteButton = new TaskModuleAction(Strings.ManageFavorites, new { data = JsonConvert.SerializeObject(new AdaptiveTaskModuleCardAction { Text = BotCommands.ShowFavoriteTaskModule, ActivityReferenceId = activityReferenceId }) });
            card.Buttons.Add(addFavoriteButton);

            CardAction otherRoomButton = new TaskModuleAction(Strings.OtherRooms, new { data = JsonConvert.SerializeObject(new AdaptiveTaskModuleCardAction { Text = BotCommands.ShowOtherRoomsTaskModule, ActivityReferenceId = activityReferenceId }) });
            card.Buttons.Add(otherRoomButton);

            if (rooms?.Schedules.Count > 0)
            {
                CardAction refreshListButton = new CardAction
                {
                    Title = Strings.Refresh,
                    Type = ActionTypes.MessageBack,
                    Text = BotCommands.RefreshList,
                    DisplayText = string.Empty,
                    Value = JsonConvert.SerializeObject(new AdaptiveTaskModuleCardAction { Text = BotCommands.RefreshList, ActivityReferenceId = activityReferenceId }),
                };

                card.Buttons.Add(refreshListButton);
            }

            var attachment = new Attachment()
            {
                ContentType = "application/vnd.microsoft.teams.card.list",
                Content = card,
            };
            return attachment;
        }
    }
}

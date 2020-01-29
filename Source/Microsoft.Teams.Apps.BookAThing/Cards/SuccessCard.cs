// <copyright file="SuccessCard.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.BookAThing.Models;
    using Microsoft.Teams.Apps.BookAThing.Resources;
    using Newtonsoft.Json;

    /// <summary>
    /// Class having method to get success attachment once meeting creation is successful.
    /// </summary>
    public static class SuccessCard
    {
        /// <summary>
        /// Get success card after meeting creation.
        /// </summary>
        /// <param name="meeting">Meeting model containig meeting details which needs to be display to user.</param>
        /// <param name="timeZone">User local time zone.</param>
        /// <returns>Adaptive card attachment indicating successful meeting creation.</returns>
        public static Attachment GetSuccessAttachment(MeetingViewModel meeting, string timeZone)
        {
            string greenBar = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAABhCAIAAACRaPz+AAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4yMfEgaZUAAAAcSURBVDhPY5h4KBAZjfJH+aP8Uf4ofyTyDwUCAAZTG+Jp0gBvAAAAAElFTkSuQmCC";
            var startUTCDateTime = DateTime.Parse(meeting.StartDateTime, null, DateTimeStyles.RoundtripKind);
            var endUTCDateTime = DateTime.Parse(meeting.EndDateTime, null, DateTimeStyles.RoundtripKind);
            var dateString = string.Format(CultureInfo.CurrentCulture, Strings.DateFormat, "{{DATE(" + startUTCDateTime.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'", CultureInfo.InvariantCulture) + ", SHORT)}}", "{{TIME(" + startUTCDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture) + ")}}", endUTCDateTime.Subtract(startUTCDateTime).TotalMinutes);

            var cancelMeetingAction = new AdaptiveSubmitAction
            {
                Title = Strings.CancelMeeting,
                Data = new AdaptiveSubmitActionData
                {
                    Msteams = new CardAction
                    {
                        Type = ActionTypes.MessageBack,
                        Text = BotCommands.CancelMeeting,
                        DisplayText = string.Empty,
                        Value = JsonConvert.SerializeObject(new MeetingViewModel
                        {
                            MeetingId = meeting.MeetingId,
                            StartDateTime = meeting.StartDateTime,
                            EndDateTime = meeting.EndDateTime,
                            Subject = Strings.MeetingSubject,
                            RoomName = meeting.RoomName,
                            BuildingName = meeting.BuildingName,
                        }),
                    },
                },
            };

            var card = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                    {
                        new AdaptiveColumnSet
                        {
                            Columns = new List<AdaptiveColumn>
                            {
                                new AdaptiveColumn { Width = AdaptiveColumnWidth.Auto, Items = new List<AdaptiveElement> { new AdaptiveImage { Url = new Uri(greenBar), PixelWidth = 4 } } },
                                new AdaptiveColumn
                                {
                                    Width = AdaptiveColumnWidth.Stretch,
                                    Items = new List<AdaptiveElement>
                                    {
                                        new AdaptiveTextBlock { Text = Strings.MeetingBooked, Wrap = true, Size = AdaptiveTextSize.Large, Weight = AdaptiveTextWeight.Bolder },
                                        new AdaptiveTextBlock { Text = meeting.RoomName, Wrap = true, Spacing = AdaptiveSpacing.Small },
                                        new AdaptiveTextBlock { Text = meeting.BuildingName, Wrap = true, Spacing = AdaptiveSpacing.Small },
                                        new AdaptiveTextBlock { Text = dateString, Wrap = true, Spacing = AdaptiveSpacing.Small },
                                    },
                                },
                            },
                        },
                    },
                Actions = new List<AdaptiveAction>(),
            };

            if (meeting.IsFavourite)
            {
                card.Actions = new List<AdaptiveAction>
                {
                    cancelMeetingAction,
                    new AdaptiveOpenUrlAction { Title = Strings.Share, Url = new Uri(meeting.WebLink) },
                };
            }
            else
            {
                card.Actions = new List<AdaptiveAction>
                    {
                        cancelMeetingAction,
                        new AdaptiveSubmitAction
                        {
                            Title = Strings.AddFavorite,
                            Data = new AdaptiveSubmitActionData
                            {
                                Msteams = new CardAction
                                {
                                    Type = ActionTypes.MessageBack,
                                    Text = BotCommands.AddFavorite,
                                    DisplayText = string.Empty,
                                    Value = JsonConvert.SerializeObject(new MeetingViewModel
                                    {
                                        RoomEmail = meeting.RoomEmail,
                                        BuildingName = meeting.BuildingName,
                                        RoomName = meeting.RoomName,
                                        BuildingEmail = meeting.BuildingEmail,
                                    }),
                                },
                            },
                        },
                        new AdaptiveOpenUrlAction { Title = Strings.Share, Url = new Uri(meeting.WebLink) },
                    };
            }

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
            return adaptiveCardAttachment;
        }
    }
}

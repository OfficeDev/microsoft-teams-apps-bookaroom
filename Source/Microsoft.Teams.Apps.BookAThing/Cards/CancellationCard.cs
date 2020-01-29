// <copyright file="CancellationCard.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.BookAThing.Resources;

    /// <summary>
    /// Class to get attachment for cancellation confirmation once user cancels meeting.
    /// </summary>
    public static class CancellationCard
    {
        /// <summary>
        /// Get attachment indicating meeting is cancelled.
        /// </summary>
        /// <param name="selectedRoom">Selected room email.</param>
        /// <param name="buildingName">Name of building to which room is associated with.</param>
        /// <param name="startUTCDateTime">Meeting UTC start date time.</param>
        /// <param name="endUTCDateTime">Meeting UTC end date time.</param>
        /// <param name="timezone">User local time zone.</param>
        /// <returns>Adaptive card attachment for meeting cancellation confirmation.</returns>
        public static Attachment GetCancellationAttachment(string selectedRoom, string buildingName, DateTime startUTCDateTime, DateTime endUTCDateTime, string timezone)
        {
            string redBar = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAABhCAIAAACRaPz+AAAABGdBTUEAALGPC/xhBQAAAAlwSFlzAAAOwgAADsIBFShKgAAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMC4yMfEgaZUAAAAbSURBVDhPYzhi6I2MRvmj/FH+KH+UPxL5ht4ACQ5eX5Tqr7oAAAAASUVORK5CYII=";
            var dateString = string.Format(CultureInfo.CurrentCulture, Strings.DateFormat, "{{DATE(" + startUTCDateTime.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'", CultureInfo.InvariantCulture) + ", SHORT)}}", "{{TIME(" + startUTCDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture) + ")}}", endUTCDateTime.Subtract(startUTCDateTime).TotalMinutes);

            AdaptiveCard card = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn { Width = AdaptiveColumnWidth.Auto, Items = new List<AdaptiveElement> { new AdaptiveImage { Url = new Uri(redBar) } } },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock { Text = Strings.MeetingCancelled, Wrap = true, Size = AdaptiveTextSize.Large, Weight = AdaptiveTextWeight.Bolder },
                                    new AdaptiveTextBlock { Text = selectedRoom, Wrap = true, Spacing = AdaptiveSpacing.Small },
                                    new AdaptiveTextBlock { Text = buildingName, Wrap = true, Spacing = AdaptiveSpacing.Small },
                                    new AdaptiveTextBlock { Text = dateString, Wrap = true, Spacing = AdaptiveSpacing.Small },
                                },
                            },
                        },
                    },
                },
            };

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
            return adaptiveCardAttachment;
        }
    }
}

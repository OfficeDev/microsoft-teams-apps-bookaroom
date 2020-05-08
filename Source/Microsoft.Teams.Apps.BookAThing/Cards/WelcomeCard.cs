// <copyright file="WelcomeCard.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Cards
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.BookAThing.Constants;
    using Microsoft.Teams.Apps.BookAThing.Models;
    using Microsoft.Teams.Apps.BookAThing.Resources;

    /// <summary>
    /// Class having method to return welcome card attachment.
    /// </summary>
    public static class WelcomeCard
    {
        /// <summary>
        /// Get welcome card attachment.
        /// </summary>
        /// <param name="welcomeCardImageUrl">Welcome card image URL.</param>
        /// <returns>Adaptive card attachment for bot introduction and bot commands to start with.</returns>
        public static Attachment GetWelcomeCardAttachment(Uri welcomeCardImageUrl)
        {
            AdaptiveCard card = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = welcomeCardImageUrl,
                                        Size = AdaptiveImageSize.Large,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Large,
                                        Wrap = true,
                                        Text = Strings.WelcomeCardTitle,
                                        Weight = AdaptiveTextWeight.Bolder,
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Default,
                                        Wrap = true,
                                        Text = Strings.WelcomeCardContent,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = Strings.WelcomeCardContentPart1,
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = Strings.WelcomeCardContentPart2,
                        Wrap = true,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = Strings.BookRoom,
                        Data = new AdaptiveSubmitActionData
                        {
                            Msteams = new CardAction
                            {
                                Type = ActionTypes.MessageBack,
                                Text = BotCommands.BookAMeeting,
                            },
                        },
                    },
                    new AdaptiveSubmitAction
                    {
                        Title = Strings.ManageFavorites,
                        Data = new AdaptiveSubmitActionData
                        {
                            Msteams = new CardAction
                            {
                                Type = ActionTypes.MessageBack,
                                Text = BotCommands.ManageFavorites,
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

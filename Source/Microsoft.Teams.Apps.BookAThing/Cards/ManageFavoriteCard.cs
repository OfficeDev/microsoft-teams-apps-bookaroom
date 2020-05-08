// <copyright file="ManageFavoriteCard.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Cards
{
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.BookAThing.Constants;
    using Microsoft.Teams.Apps.BookAThing.Models;
    using Microsoft.Teams.Apps.BookAThing.Resources;
    using Newtonsoft.Json;

    /// <summary>
    /// Class having method to return manage favorites card attachment from which task module can be invoked.
    /// </summary>
    public static class ManageFavoriteCard
    {
        /// <summary>
        /// Get manage favorite card attachment.
        /// </summary>
        /// <param name="appId"> Microsoft app id.</param>
        /// <returns>An attachment.</returns>
        public static Attachment GetManageFavoriteAttachment(string appId)
        {
            var card = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock { Text = Strings.ManageFavoriteCardDescription, Wrap = true },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = Strings.ManageFavorites,
                        Data = new AdaptiveSubmitActionData
                        {
                            Msteams = new TaskModuleAction(Strings.ManageFavorites, new AdaptiveTaskModuleCardAction(appId) { Text = BotCommands.ShowFavoriteTaskModule }),
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

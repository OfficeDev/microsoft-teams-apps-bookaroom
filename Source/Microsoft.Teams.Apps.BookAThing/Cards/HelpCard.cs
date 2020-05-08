// <copyright file="HelpCard.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Cards
{
    using System.Collections.Generic;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.BookAThing.Constants;
    using Microsoft.Teams.Apps.BookAThing.Models;
    using Microsoft.Teams.Apps.BookAThing.Resources;
    using Newtonsoft.Json;

    /// <summary>
    /// Class having method for returning help commands card attachment.
    /// </summary>
    public static class HelpCard
    {
        /// <summary>
        /// Get help card attachment.
        /// </summary>
        /// <param name="appId"> Microsoft app id.</param>
        /// <returns>List of attachments.</returns>
        public static List<Attachment> GetHelpAttachments(string appId)
        {
            List<Attachment> attachments = new List<Attachment>();

            List<CardAction> buttons = new List<CardAction>();
            buttons.AddRange(new List<CardAction>
            {
                new CardAction(ActionTypes.MessageBack, Strings.BookRoom, text: BotCommands.BookAMeeting, displayText: Strings.BookRoom),
                new TaskModuleAction(Strings.ManageFavorites, new AdaptiveTaskModuleCardAction(appId) { Text = BotCommands.ShowFavoriteTaskModule }),
            });

            var heroCard = new HeroCard
            {
                Text = Strings.SupportedCommands,
                Buttons = buttons,
            };

            attachments.Add(heroCard.ToAttachment());
            return attachments;
        }
    }
}

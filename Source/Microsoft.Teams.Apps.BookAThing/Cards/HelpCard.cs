// <copyright file="HelpCard.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Cards
{
    using System.Collections.Generic;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
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
        /// <returns>List of attachments.</returns>
        public static List<Attachment> GetHelpAttachments()
        {
            List<Attachment> attachments = new List<Attachment>();

            List<CardAction> buttons = new List<CardAction>();
            buttons.AddRange(new List<CardAction>
            {
                new CardAction(ActionTypes.MessageBack, Strings.BookRoom, text: BotCommands.BookAMeeting, displayText: Strings.BookRoom, value: string.Empty),
                new TaskModuleAction(Strings.ManageFavorites, new { data = JsonConvert.SerializeObject(new AdaptiveTaskModuleCardAction { Text = BotCommands.ShowFavoriteTaskModule }) }),
            });

            var heroCard = new HeroCard
            {
                Buttons = buttons,
            };

            attachments.Add(heroCard.ToAttachment());
            return attachments;
        }
    }
}

// <copyright file="ListCard.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Models
{
    using System.Collections.Generic;
    using Microsoft.Bot.Schema;
    using Newtonsoft.Json;

    /// <summary>
    /// List card root class.
    /// </summary>
    public class ListCard
    {
        /// <summary>
        /// Gets or sets title of card.
        /// </summary>
        [JsonProperty("title")]
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets list items.
        /// </summary>
        [JsonProperty("items")]
        public List<ListItem> Items { get; set; }

        /// <summary>
        /// Gets or sets buttons.
        /// </summary>
        [JsonProperty("buttons")]
        public List<CardAction> Buttons { get; set; }
    }
}

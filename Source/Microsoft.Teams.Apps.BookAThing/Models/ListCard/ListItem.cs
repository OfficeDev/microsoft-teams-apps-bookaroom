// <copyright file="ListItem.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Models
{
    using Microsoft.Bot.Schema;
    using Newtonsoft.Json;

    /// <summary>
    /// List card Item class.
    /// </summary>
    public class ListItem
    {
        /// <summary>
        /// Gets or sets type of item.
        /// </summary>
        [JsonProperty("type")]
        public string Type { get; set; }

        /// <summary>
        /// Gets or sets id.
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets title of item.
        /// </summary>
        [JsonProperty("title")]
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets subtitle of item.
        /// </summary>
        [JsonProperty("subtitle")]
        public string Subtitle { get; set; }

        /// <summary>
        /// Gets or sets tap action.
        /// </summary>
        [JsonProperty("tap")]
        public CardAction Tap { get; set; }

        /// <summary>
        /// Gets or sets icon.
        /// </summary>
        [JsonProperty("icon")]
        public string Icon { get; set; }
    }
}

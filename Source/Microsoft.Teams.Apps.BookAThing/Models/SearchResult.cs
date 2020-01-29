// <copyright file="SearchResult.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Models
{
    using Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities;
    using Newtonsoft.Json;

    /// <summary>
    /// Class containing properties required for room search dropdown in client app.
    /// </summary>
    public class SearchResult : UserFavoriteRoomEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SearchResult"/> class.
        /// </summary>
        /// <param name="meetingRoomCollection">Rooms collection table entity class.</param>
        public SearchResult(MeetingRoomEntity meetingRoomCollection)
        {
            this.RoomName = meetingRoomCollection?.RoomName;
            this.BuildingName = meetingRoomCollection?.BuildingName;
            this.RowKey = meetingRoomCollection?.RowKey;
            this.PartitionKey = meetingRoomCollection?.PartitionKey;
        }

        /// <summary>
        /// Gets or sets room name.
        /// </summary>
        [JsonProperty("label")]
        public string Label { get; set; }

        /// <summary>
        /// Gets or sets room email.
        /// </summary>
        [JsonProperty("value")]
        public string Value { get; set; }

        /// <summary>
        /// Gets or sets building name.
        /// </summary>
        [JsonProperty("sublabel")]
        public string Sublabel { get; set; }

        /// <summary>
        /// Gets or sets room availability status.
        /// </summary>
        public string Status { get; set; }
    }
}

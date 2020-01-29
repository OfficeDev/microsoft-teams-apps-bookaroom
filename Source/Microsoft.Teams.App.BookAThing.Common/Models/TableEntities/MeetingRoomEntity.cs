// <copyright file="MeetingRoomEntity.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities
{
    using System.ComponentModel.DataAnnotations;
    using Microsoft.Azure.Search;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Table to store rooms and buildings information.
    /// </summary>
    public class MeetingRoomEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets email address associated with building.
        /// </summary>
        public string BuildingEmail
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        /// Gets or sets email address associated with room.
        /// </summary>
        public string RoomEmail
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets Key which will be used by azure search indexer. Here key is room id received from Graph API.
        /// </summary>
        [Key]
        public string Key { get; set; }

        /// <summary>
        /// Gets or sets name of room.
        /// </summary>
        [IsSearchable]
        [IsSortable]
        [IsFilterable]
        public string RoomName { get; set; }

        /// <summary>
        /// Gets or sets name of building.
        /// </summary>
        [IsSearchable]
        [IsFilterable]
        public string BuildingName { get; set; }
    }
}

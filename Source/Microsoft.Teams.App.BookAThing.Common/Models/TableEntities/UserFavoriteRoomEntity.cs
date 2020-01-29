// <copyright file="UserFavoriteRoomEntity.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// User favorites table entity class.
    /// </summary>
    public class UserFavoriteRoomEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets user Active Directory object Id of user.
        /// </summary>
        public string UserAdObjectId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        /// Gets or sets email associated with room.
        /// </summary>
        public string RoomEmail
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets room name.
        /// </summary>
        public string RoomName { get; set; }

        /// <summary>
        /// Gets or sets building name.
        /// </summary>
        public string BuildingName { get; set; }

        /// <summary>
        /// Gets or sets email address associated with building.
        /// </summary>
        public string BuildingEmail { get; set; }
    }
}

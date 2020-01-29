// <copyright file="UserFavoriteRooms.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Models
{
    using System.Collections.Generic;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities;

    /// <summary>
    /// Favorite room collection for user.
    /// </summary>
    public class UserFavoriteRooms
    {
        /// <summary>
        /// Gets or sets list of rooms.
        /// </summary>
        public List<UserFavoriteRoomEntity> Rooms { get; set; }

        /// <summary>
        /// Gets or sets Active Directory object Id.
        /// </summary>
        public string UserObjectIdentifier { get; set; }
    }
}

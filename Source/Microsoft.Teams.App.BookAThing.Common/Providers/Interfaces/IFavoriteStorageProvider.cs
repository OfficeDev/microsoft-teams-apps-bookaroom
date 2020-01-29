// <copyright file="IFavoriteStorageProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Providers.Storage
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities;

    /// <summary>
    /// Storage provider for insert, update and delete operation on FavoriteRooms table.
    /// </summary>
    public interface IFavoriteStorageProvider
    {
        /// <summary>
        /// Add room to user favorite.
        /// </summary>
        /// <param name="room">Room entity object.</param>
        /// <returns>List of favorite rooms for user after adding a room.</returns>
        Task<IList<UserFavoriteRoomEntity>> AddAsync(UserFavoriteRoomEntity room);

        /// <summary>
        /// Adds rooms to favorite for user.
        /// </summary>
        /// <param name="rooms">List of favorite rooms.</param>
        /// <returns>Boolean indicating operation result.</returns>
        Task<bool> AddBatchAsync(IList<UserFavoriteRoomEntity> rooms);

        /// <summary>
        /// Removes all favorite rooms of user.
        /// </summary>
        /// <param name="userIdentifier">User object identifier.</param>
        /// <returns>Boolean indicating operation result.</returns>
        Task<bool> DeleteAllAsync(string userIdentifier);

        /// <summary>
        /// Delete favorites rooms added of building.
        /// </summary>
        /// <param name="roomEmails">List of room email.</param>
        /// <param name="buildingEmail">Building email.</param>
        /// <returns>Boolean indicating operation result.</returns>
        Task<bool> DeleteAsync(IList<string> roomEmails, string buildingEmail);

        /// <summary>
        /// Get user favorite rooms.
        /// </summary>
        /// <param name="userIdentifier">User object identifier.</param>
        /// <param name="rowKey">Room email id.</param>
        /// <returns>List of favorite rooms.</returns>
        Task<IList<UserFavoriteRoomEntity>> GetAsync(string userIdentifier, string rowKey = null);
    }
}
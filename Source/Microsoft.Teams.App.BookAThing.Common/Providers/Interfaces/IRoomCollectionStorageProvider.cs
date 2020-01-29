// <copyright file="IRoomCollectionStorageProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Providers.Storage
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities;

    /// <summary>
    /// Storage provider for insert, update and delete operation on RoomCollection table.
    /// </summary>
    public interface IRoomCollectionStorageProvider
    {
        /// <summary>
        /// Add rooms to storage.
        /// </summary>
        /// <param name="rooms">List of rooms.</param>
        /// <returns>Boolean indicating operation result.</returns>
        Task<bool> AddAsync(IList<MeetingRoomEntity> rooms);

        /// <summary>
        /// Delete all rooms associated to a building.
        /// </summary>
        /// <param name="rooms">List of rooms.</param>
        /// <returns>Boolean indicating operation result.</returns>
        Task<bool> DeleteAsync(IList<MeetingRoomEntity> rooms);

        /// <summary>
        /// Get all rooms stored in azure table storage.
        /// </summary>
        /// <returns>List of all rooms.</returns>
        Task<IList<MeetingRoomEntity>> GetAllAsync();

        /// <summary>
        /// Get all rooms associated with a building.
        /// </summary>
        /// <param name="buildingEmail">Building alias.</param>
        /// <returns>List of rooms associated with building.</returns>
        Task<IList<MeetingRoomEntity>> GetAsync(string buildingEmail);

        /// <summary>
        /// Get 'N' rooms from storage where 'N' is room count.
        /// </summary>
        /// <param name="roomCount">Number of rooms to be fetched.</param>
        /// <returns>List of meeting rooms.</returns>
        Task<IList<MeetingRoomEntity>> GetNRoomsAsync(int roomCount);
    }
}
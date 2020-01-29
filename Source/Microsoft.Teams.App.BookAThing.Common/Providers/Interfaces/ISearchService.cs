// <copyright file="ISearchService.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities;

    /// <summary>
    /// SearchService which will help in creating index, indexer and data source if it doesn't exists
    /// for indexing table which will be used for search by bot.
    /// </summary>
    public interface ISearchService
    {
        /// <summary>
        /// Search room or building by name.
        /// </summary>
        /// <param name="searchQuery">Search string.</param>
        /// <returns>List of rooms.</returns>
        Task<IList<MeetingRoomEntity>> SearchRoomsAsync(string searchQuery);

        /// <summary>
        /// Create index, indexer and data source if doesn't exists.
        /// </summary>
        /// <returns>Tracking task.</returns>
        Task InitializeAsync();
    }
}
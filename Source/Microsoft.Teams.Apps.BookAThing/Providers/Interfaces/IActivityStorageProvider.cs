// <copyright file="IActivityStorageProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Providers.Storage
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.BookAThing.Models.TableEntities;

    /// <summary>
    /// Storage provider for fetch, insert and update operation on Activities table.
    /// </summary>
    public interface IActivityStorageProvider
    {
        /// <summary>
        /// Add or update activity id.
        /// </summary>
        /// <param name="activity">Activity table entity.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<bool> AddAsync(ActivityEntity activity);

        /// <summary>
        /// Get activity Ids.
        /// </summary>
        /// <param name="userObjectIdentifer">Active Directory object Id of user.</param>
        /// <param name="activityReferenceId">Unique GUID referencing to activity id.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<ActivityEntity> GetAsync(string userObjectIdentifer, string activityReferenceId);
    }
}
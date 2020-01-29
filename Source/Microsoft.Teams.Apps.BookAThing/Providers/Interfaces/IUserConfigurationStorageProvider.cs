// <copyright file="IUserConfigurationStorageProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Providers.Storage
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.BookAThing.Models.TableEntities;

    /// <summary>
    /// Storage provider for fetch, insert and update operation on UserConfiguration table.
    /// </summary>
    public interface IUserConfigurationStorageProvider
    {
        /// <summary>
        /// Add or update user configuration.
        /// </summary>
        /// <param name="userConfiguration">User configuration entity.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<bool> AddAsync(UserConfigurationEntity userConfiguration);

        /// <summary>
        /// Get user configuration.
        /// </summary>
        /// <param name="userObjectIdentifer">Active Directory object Id of user.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<UserConfigurationEntity> GetAsync(string userObjectIdentifer);
    }
}
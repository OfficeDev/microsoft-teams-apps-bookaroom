// <copyright file="IExchangeSyncHelper.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.SyncService.Service
{
    using System.Threading.Tasks;

    /// <summary>
    /// Methods for performing exchange to table storage sync operation.
    /// </summary>
    public interface IExchangeSyncHelper
    {
        /// <summary>
        /// Process exchange to storage sync.
        /// </summary>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task ExchangeToStorageExportAsync();
    }
}
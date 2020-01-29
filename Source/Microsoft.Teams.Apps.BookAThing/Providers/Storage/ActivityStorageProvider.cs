// <copyright file="ActivityStorageProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Providers.Storage
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.Teams.Apps.BookAThing.Models.TableEntities;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Storage provider for fetch, insert and update operation on Activities table.
    /// </summary>
    public class ActivityStorageProvider : IActivityStorageProvider
    {
        /// <summary>
        /// Table name in Azure table storage.
        /// </summary>
        private const string TableName = "ActivityEntities";

        /// <summary>
        /// Telemetry client for logging events and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Task for initialization.
        /// </summary>
        private readonly Lazy<Task> initializeTask;

        /// <summary>
        /// Provides a service client for accessing the Microsoft Azure Table service.
        /// </summary>
        private CloudTableClient cloudTableClient;

        /// <summary>
        /// Represents a table in the Microsoft Azure Table service.
        /// </summary>
        private CloudTable cloudTable;

        /// <summary>
        /// Initializes a new instance of the <see cref="ActivityStorageProvider"/> class.
        /// </summary>
        /// <param name="storageConnectionString">Azure Table Storage connection string.</param>
        /// <param name="telemetryClient">Telemetry client for logging events and errors.</param>
        public ActivityStorageProvider(string storageConnectionString, TelemetryClient telemetryClient)
        {
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(storageConnectionString));
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Get activity ids.
        /// </summary>
        /// <param name="userObjectIdentifer">Active Directory object Id of user.</param>
        /// <param name="activityReferenceId">Unique GUID referencing to activity id.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task<ActivityEntity> GetAsync(string userObjectIdentifer, string activityReferenceId)
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                var retrieveOperation = TableOperation.Retrieve<ActivityEntity>(userObjectIdentifer, activityReferenceId);
                var room = await this.cloudTable.ExecuteAsync(retrieveOperation).ConfigureAwait(false);

                return (ActivityEntity)room.Result;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Add or update activity id.
        /// </summary>
        /// <param name="activity">Activity table entity.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task<bool> AddAsync(ActivityEntity activity)
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                TableOperation insertOrMergeOperation = TableOperation.InsertOrReplace(activity);
                TableResult result = await this.cloudTable.ExecuteAsync(insertOrMergeOperation).ConfigureAwait(false);
                return result?.Result != null;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Ensure table storage connection is initialized.
        /// </summary>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task EnsureInitializedAsync()
        {
            await this.initializeTask.Value;
        }

        /// <summary>
        /// Create tables if it doesn't exists.
        /// </summary>
        /// <param name="connectionString">Storage account connection string.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task InitializeAsync(string connectionString)
        {
            try
            {
                var storageAccount = CloudStorageAccount.Parse(connectionString);
                this.cloudTableClient = storageAccount.CreateCloudTableClient();
                this.cloudTable = this.cloudTableClient.GetTableReference(TableName);
                await this.cloudTable.CreateIfNotExistsAsync().ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }
    }
}

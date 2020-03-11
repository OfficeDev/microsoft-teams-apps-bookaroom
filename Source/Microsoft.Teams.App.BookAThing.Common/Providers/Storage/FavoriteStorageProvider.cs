// <copyright file="FavoriteStorageProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Providers.Storage
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.RetryPolicies;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Storage provider for insert, update and delete operation on FavoriteRooms table.
    /// </summary>
    public class FavoriteStorageProvider : IFavoriteStorageProvider
    {
        /// <summary>
        /// Max number of rooms for a batch operation.
        /// </summary>
        private const int RoomsPerBatch = 100;

        /// <summary>
        /// Table name in Azure table storage.
        /// </summary>
        private const string TableName = "FavoriteRooms";

        /// <summary>
        /// Building email column name in table.
        /// </summary>
        private const string BuildingEmailColumnName = "BuildingEmail";

        /// <summary>
        /// Task for initialization.
        /// </summary>
        private readonly Lazy<Task> initializeTask;

        /// <summary>
        /// Telemetry client for logging events and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Provides a service client for accessing the Microsoft Azure Table service.
        /// </summary>
        private CloudTableClient cloudTableClient;

        /// <summary>
        /// Represents a table in the Microsoft Azure Table service.
        /// </summary>
        private CloudTable cloudTable;

        /// <summary>
        /// Initializes a new instance of the <see cref="FavoriteStorageProvider"/> class.
        /// </summary>
        /// <param name="connectionString">Table storage connection string.</param>
        /// <param name="telemetryClient">Telemetry client for logging events and errors.</param>
        public FavoriteStorageProvider(string connectionString, TelemetryClient telemetryClient)
        {
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(connectionString));
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Enum indicating batch operation type.
        /// </summary>
        private enum BatchOperation
        {
            Insert,
            Delete,
        }

        /// <summary>
        /// Get user favorite rooms.
        /// </summary>
        /// <param name="userObjectIdentifier">Active Directory object Id of user.</param>
        /// <param name="roomEmail">Room email id.</param>
        /// <returns>List of favorite rooms for user.</returns>
        public async Task<IList<UserFavoriteRoomEntity>> GetAsync(string userObjectIdentifier, string roomEmail = null)
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                string partitionCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, userObjectIdentifier);
                TableQuery<UserFavoriteRoomEntity> query;

                if (!string.IsNullOrEmpty(roomEmail))
                {
                    string rowKeyCondition = TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, roomEmail);
                    query = new TableQuery<UserFavoriteRoomEntity>().Where(TableQuery.CombineFilters(partitionCondition, TableOperators.And, rowKeyCondition));
                }
                else
                {
                    query = new TableQuery<UserFavoriteRoomEntity>().Where(partitionCondition);
                }

                TableContinuationToken continuationToken = null;
                var rooms = new List<UserFavoriteRoomEntity>();

                do
                {
                    var queryResult = await this.cloudTable.ExecuteQuerySegmentedAsync(query, continuationToken).ConfigureAwait(false);
                    rooms.AddRange(queryResult?.Results);
                    continuationToken = queryResult?.ContinuationToken;
                }
                while (continuationToken != null);

                return rooms;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Add room to user favorite.
        /// </summary>
        /// <param name="room">Room entity object.</param>
        /// <returns>List of favorite rooms for user after adding a room.</returns>
        public async Task<IList<UserFavoriteRoomEntity>> AddAsync(UserFavoriteRoomEntity room)
        {
            try
            {
                await this.EnsureInitializedAsync();
                var insertOrMergeOperation = TableOperation.InsertOrReplace(room);
                await this.cloudTable.ExecuteAsync(insertOrMergeOperation).ConfigureAwait(false);
                return await this.GetAsync(room.PartitionKey).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Adds rooms to favorite for user.
        /// </summary>
        /// <param name="rooms">List of favorite rooms.</param>
        /// <returns>Returns true if batch operation for inserting favorite rooms for user succeeds.</returns>
        public async Task<bool> AddBatchAsync(IList<UserFavoriteRoomEntity> rooms)
        {
            try
            {
                await this.EnsureInitializedAsync();
                return await this.ExecuteBatchOperationAsync(BatchOperation.Insert, rooms).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Removes all favorite rooms of user.
        /// </summary>
        /// <param name="userObjectIdentifier">Active Directory object id of user.</param>
        /// <returns>Returns true if batch operation for deleting favorite rooms of user succeeds.</returns>
        public async Task<bool> DeleteAllAsync(string userObjectIdentifier)
        {
            try
            {
                await this.EnsureInitializedAsync();
                var favoriteRooms = await this.GetAsync(userObjectIdentifier).ConfigureAwait(false);

                if (favoriteRooms?.Count > 0)
                {
                    return await this.ExecuteBatchOperationAsync(BatchOperation.Delete, favoriteRooms).ConfigureAwait(false);
                }

                return true;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Executes batch add or delete operation on Azure table storage.
        /// </summary>
        /// <param name="batchOperation">Batch operation to be performed.</param>
        /// <param name="rooms">List of rooms.</param>
        /// <returns>Returns true if batch operation is successful else throws exception for error.</returns>
        private async Task<bool> ExecuteBatchOperationAsync(BatchOperation batchOperation, IList<UserFavoriteRoomEntity> rooms)
        {
            var tableBatchOperation = new TableBatchOperation();
            try
            {
                int batchCount = (int)Math.Ceiling((double)rooms.Count / RoomsPerBatch);

                for (int i = 0; i < batchCount; i++)
                {
                    tableBatchOperation.Clear();
                    var roomsBatch = rooms.Skip(i * RoomsPerBatch).Take(RoomsPerBatch);
                    foreach (var room in roomsBatch)
                    {
                        tableBatchOperation.Add(batchOperation == BatchOperation.Insert ? TableOperation.Insert(room) : TableOperation.Delete(room));
                    }

                    if (tableBatchOperation.Count > 0)
                    {
                        await this.cloudTable.ExecuteBatchAsync(tableBatchOperation).ConfigureAwait(false);
                    }
                }

                return true;
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
        /// Create Azure storage table if it doesn't exists.
        /// </summary>
        /// <param name="connectionString">Storage account connection string.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task InitializeAsync(string connectionString)
        {
            // Exponential retry policy with backoff of 3 seconds and 5 retries.
            var exponentialRetryPolicy = new TableRequestOptions()
            {
                RetryPolicy = new ExponentialRetry(TimeSpan.FromSeconds(3), 5),
            };

            try
            {
                var storageAccount = CloudStorageAccount.Parse(connectionString);
                this.cloudTableClient = storageAccount.CreateCloudTableClient();
                this.cloudTableClient.DefaultRequestOptions = exponentialRetryPolicy;
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

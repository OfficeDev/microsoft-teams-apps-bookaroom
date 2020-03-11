// <copyright file="RoomCollectionStorageProvider.cs" company="Microsoft Corporation">
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
    /// Storage provider for fetch, insert, update and delete operation on RoomCollection table.
    /// </summary>
    public class RoomCollectionStorageProvider : IRoomCollectionStorageProvider
    {
        /// <summary>
        /// Table name in Azure table storage.
        /// </summary>
        public const string TableName = "RoomCollection";

        /// <summary>
        /// Max number of rooms for a batch operation.
        /// </summary>
        private const int RoomsPerBatch = 100;

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
        /// Initializes a new instance of the <see cref="RoomCollectionStorageProvider"/> class.
        /// </summary>
        /// <param name="connectionString">Table storage connection string.</param>
        /// <param name="telemetryClient">Telemetry client for logging events and errors.</param>
        public RoomCollectionStorageProvider(string connectionString, TelemetryClient telemetryClient)
        {
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(connectionString));
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Enum indicating batch operation type.
        /// </summary>
        private enum BatchOperation
        {
            InsertOrReplace,
            Delete,
        }

        /// <summary>
        /// Get all rooms associated with a building.
        /// </summary>
        /// <param name="buildingEmail">Building alias.</param>
        /// <returns>List of rooms associated with building.</returns>
        public async Task<IList<MeetingRoomEntity>> GetAsync(string buildingEmail)
        {
            try
            {
                await this.EnsureInitializedAsync();
                string partitionKeyCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, buildingEmail);
                var query = new TableQuery<MeetingRoomEntity>().Where(partitionKeyCondition);
                TableContinuationToken continuationToken = null;
                var rooms = new List<MeetingRoomEntity>();

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
        /// Get top 'N' rooms from Azure table storage where 'N' is room count.
        /// </summary>
        /// <param name="roomCount">Number of rooms to be fetched.</param>
        /// <returns>List of meeting rooms.</returns>
        public async Task<IList<MeetingRoomEntity>> GetNRoomsAsync(int roomCount)
        {
            try
            {
                await this.EnsureInitializedAsync();
                var query = new TableQuery<MeetingRoomEntity>
                {
                    TakeCount = roomCount,
                    FilterString = "IsDeleted eq false",
                };

                TableContinuationToken continuationToken = null;
                var rooms = new List<MeetingRoomEntity>();
                var queryResult = await this.cloudTable.ExecuteQuerySegmentedAsync(query, continuationToken).ConfigureAwait(false);
                if (queryResult?.Results != null)
                {
                    rooms.AddRange(queryResult?.Results);
                }

                return rooms;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Delete all rooms associated with a building.
        /// </summary>
        /// <param name="rooms">List of rooms.</param>
        /// <returns>Returns true if batch operation for deleting rooms succeeds.</returns>
        public async Task<bool> DeleteAsync(IList<MeetingRoomEntity> rooms)
        {
            try
            {
                await this.EnsureInitializedAsync();
                return await this.ExecuteBatchOperationAsync(BatchOperation.Delete, rooms).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Update rooms in Azure table storage as per change in Microsoft Exchange.
        /// </summary>
        /// <param name="roomCollection">List of rooms which got deleted from Exchange.</param>
        /// <returns>Returns true if batch operation for updating rooms succeeds.</returns>
        public async Task<bool> UpdateDeletedRoomsAsync(IList<MeetingRoomEntity> roomCollection)
        {
            try
            {
                await this.EnsureInitializedAsync();
                return await this.ExecuteBatchOperationAsync(BatchOperation.InsertOrReplace, roomCollection).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Add rooms to Azure table storage.
        /// </summary>
        /// <param name="rooms">List of rooms.</param>
        /// <returns>Returns true if batch operation for inserting rooms succeeds.</returns>
        public async Task<bool> AddOrReplaceAsync(IList<MeetingRoomEntity> rooms)
        {
            try
            {
                await this.EnsureInitializedAsync();
                return await this.ExecuteBatchOperationAsync(BatchOperation.InsertOrReplace, rooms).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Get all rooms stored in Azure table storage.
        /// </summary>
        /// <returns>List of all rooms.</returns>
        public async Task<IList<MeetingRoomEntity>> GetAllAsync()
        {
            try
            {
                await this.EnsureInitializedAsync();
                var query = new TableQuery<MeetingRoomEntity>();
                TableContinuationToken continuationToken = null;
                var rooms = new List<MeetingRoomEntity>();
                do
                {
                    var queryResult = await this.cloudTable.ExecuteQuerySegmentedAsync(query, continuationToken).ConfigureAwait(false);
                    if (queryResult != null)
                    {
                        rooms.AddRange(queryResult);
                    }

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
        /// Executes batch add or delete operation on Azure table storage.
        /// </summary>
        /// <param name="batchOperation">Batch operation to be performed.</param>
        /// <param name="rooms">List of rooms.</param>
        /// <returns>Returns true if batch operation is successful else throws exception for error.</returns>
        private async Task<bool> ExecuteBatchOperationAsync(BatchOperation batchOperation, IList<MeetingRoomEntity> rooms)
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
                        tableBatchOperation.Add(batchOperation == BatchOperation.InsertOrReplace ? TableOperation.InsertOrReplace(room) : TableOperation.Delete(room));
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
        /// Ensure Azure table storage connection is initialized.
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

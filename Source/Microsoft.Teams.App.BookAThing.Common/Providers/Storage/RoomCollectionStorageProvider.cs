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
    using Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities;
    using Microsoft.WindowsAzure.Storage;
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
            Insert,
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
                TableQuery<MeetingRoomEntity> query = new TableQuery<MeetingRoomEntity>().Where(partitionKeyCondition);
                TableContinuationToken continuationToken = null;
                List<MeetingRoomEntity> rooms = new List<MeetingRoomEntity>();

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
        /// Get 'N' rooms from storage where 'N' is room count.
        /// </summary>
        /// <param name="roomCount">Number of rooms to be fetched.</param>
        /// <returns>List of meeting rooms.</returns>
        public async Task<IList<MeetingRoomEntity>> GetNRoomsAsync(int roomCount)
        {
            try
            {
                await this.EnsureInitializedAsync();
                TableQuery<MeetingRoomEntity> query = new TableQuery<MeetingRoomEntity>
                {
                    TakeCount = roomCount,
                };

                TableContinuationToken continuationToken = null;
                List<MeetingRoomEntity> rooms = new List<MeetingRoomEntity>();
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
        /// <returns>Boolean indicating operation result.</returns>
        public async Task<bool> DeleteAsync(IList<MeetingRoomEntity> rooms)
        {
            try
            {
                await this.EnsureInitializedAsync();
                TableBatchOperation deleteBatchOperation = new TableBatchOperation();
                return await this.ExecuteBatchOperation(BatchOperation.Delete, rooms);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Add rooms to storage.
        /// </summary>
        /// <param name="rooms">List of rooms.</param>
        /// <returns>Boolean indicating operation result.</returns>
        public async Task<bool> AddAsync(IList<MeetingRoomEntity> rooms)
        {
            try
            {
                await this.EnsureInitializedAsync();
                TableBatchOperation addBatchOperation = new TableBatchOperation();
                return await this.ExecuteBatchOperation(BatchOperation.Insert, rooms);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Get all rooms stored in azure table storage.
        /// </summary>
        /// <returns>List of all rooms.</returns>
        public async Task<IList<MeetingRoomEntity>> GetAllAsync()
        {
            try
            {
                await this.EnsureInitializedAsync();
                TableQuery<MeetingRoomEntity> query = new TableQuery<MeetingRoomEntity>();
                TableContinuationToken continuationToken = null;
                List<MeetingRoomEntity> rooms = new List<MeetingRoomEntity>();
                do
                {
                    var queryResult = await this.cloudTable.ExecuteQuerySegmentedAsync(query, continuationToken).ConfigureAwait(false);
                    rooms.AddRange(queryResult);
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
        /// Executes batch add or delete operation on table.
        /// </summary>
        /// <param name="batchOperation">Batch operation to be performed.</param>
        /// <param name="rooms">List of rooms.</param>
        /// <returns>Boolean indicating operation result.</returns>
        private async Task<bool> ExecuteBatchOperation(BatchOperation batchOperation, IList<MeetingRoomEntity> rooms)
        {
            try
            {
                TableBatchOperation tableBatchOperation = new TableBatchOperation();
                int count = (int)Math.Ceiling((double)rooms.Count / RoomsPerBatch);
                for (int i = 0; i < count; i++)
                {
                    var roomsBatch = rooms.Skip(i * RoomsPerBatch).Take(RoomsPerBatch);
                    foreach (var room in roomsBatch)
                    {
                        tableBatchOperation.Add(batchOperation == BatchOperation.Insert ? TableOperation.Insert(room) : TableOperation.Delete(room));
                    }
                }

                if (tableBatchOperation.Count > 0)
                {
                    var result = await this.cloudTable.ExecuteBatchAsync(tableBatchOperation).ConfigureAwait(false);
                    return result?.Count > 0;
                }

                return false;
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
        /// <param name="connectionString">storage account connection string.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task InitializeAsync(string connectionString)
        {
            try
            {
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
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

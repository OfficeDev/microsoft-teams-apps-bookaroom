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
    using Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities;
    using Microsoft.WindowsAzure.Storage;
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
                List<UserFavoriteRoomEntity> rooms = new List<UserFavoriteRoomEntity>();

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
        /// Delete favorites rooms associated with a building.
        /// </summary>
        /// <param name="roomEmails">Room emails.</param>
        /// <param name="buildingEmail">Building email.</param>
        /// <returns>Boolean indicating operation result.</returns>
        public async Task<bool> DeleteAsync(IList<string> roomEmails, string buildingEmail)
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                string buildingEmailCondition = TableQuery.GenerateFilterCondition(BuildingEmailColumnName, QueryComparisons.Equal, buildingEmail);
                TableQuery<UserFavoriteRoomEntity> query = new TableQuery<UserFavoriteRoomEntity>().Where(buildingEmailCondition);
                TableContinuationToken continuationToken = null;
                List<UserFavoriteRoomEntity> rooms = new List<UserFavoriteRoomEntity>();

                do
                {
                    var queryResult = await this.cloudTable.ExecuteQuerySegmentedAsync(query, continuationToken).ConfigureAwait(false);
                    rooms.AddRange(queryResult?.Results);
                    continuationToken = queryResult?.ContinuationToken;
                }
                while (continuationToken != null);

                var filteredRooms = rooms.Where(room => roomEmails.Contains(room.RowKey)).ToList();
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
        /// Add room to user favorite.
        /// </summary>
        /// <param name="room">Room entity object.</param>
        /// <returns>List of favorite rooms for user after adding a room.</returns>
        public async Task<IList<UserFavoriteRoomEntity>> AddAsync(UserFavoriteRoomEntity room)
        {
            try
            {
                await this.EnsureInitializedAsync();
                TableOperation insertOrMergeOperation = TableOperation.InsertOrReplace(room);
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
        /// <returns>Boolean indicating operation result.</returns>
        public async Task<bool> AddBatchAsync(IList<UserFavoriteRoomEntity> rooms)
        {
            try
            {
                await this.EnsureInitializedAsync();
                TableBatchOperation insertBatchOperation = new TableBatchOperation();
                return await this.ExecuteBatchOperation(BatchOperation.Insert, rooms);
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
        /// <returns>Boolean indicating operation result.</returns>
        public async Task<bool> DeleteAllAsync(string userObjectIdentifier)
        {
            try
            {
                await this.EnsureInitializedAsync();
                var results = await this.GetAsync(userObjectIdentifier).ConfigureAwait(false);

                if (results?.Count > 0)
                {
                    TableBatchOperation deleteBatch = new TableBatchOperation();
                    return await this.ExecuteBatchOperation(BatchOperation.Delete, results);
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
        /// Executes batch add or delete operation on table.
        /// </summary>
        /// <param name="batchOperation">Batch operation to be performed.</param>
        /// <param name="rooms">List of rooms.</param>
        /// <returns>Boolean indicating operation result.</returns>
        private async Task<bool> ExecuteBatchOperation(BatchOperation batchOperation, IList<UserFavoriteRoomEntity> rooms)
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

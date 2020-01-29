// <copyright file="SearchService.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.Azure.Search;
    using Microsoft.Azure.Search.Models;
    using Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities;
    using Microsoft.Teams.Apps.BookAThing.Common.Providers.Storage;

    /// <summary>
    /// SearchService which will help in creating index, indexer and data source if it doesn't exists
    /// for indexing table which will be used for search by bot.
    /// </summary>
    public class SearchService : ISearchService
    {
        /// <summary>
        /// Name of index to be created in Azure Search.
        /// </summary>
        private const string IndexName = "bookameeting-index";

        /// <summary>
        /// Name of indexer to be created in Azure Search.
        /// </summary>
        private const string IndexerName = "bookameeting-indexer";

        /// <summary>
        /// Name of data store to be created for Azure Search.
        /// </summary>
        private const string DataSourceName = "bookameeting-storage";

        /// <summary>
        /// Default search result count (20 results).
        /// </summary>
        private const int DefaultSearchResultCount = 20;

        /// <summary>
        /// Telemetry client for logging events and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Client that can be used to manage and query indexes and documents, as well as manage other resources, on a search service.
        /// </summary>
        private readonly SearchServiceClient searchServiceClient;

        /// <summary>
        /// Client that can be used to query an index and upload, merge, or delete documents.
        /// </summary>
        private readonly SearchIndexClient searchIndexClient;

        /// <summary>
        /// Duration gap for which indexer will run repeatedly.
        /// </summary>
        private readonly int searchIndexingIntervalInMinutes;

        /// <summary>
        /// Azure table storage connection string.
        /// </summary>
        private readonly string storageConnectionString;

        /// <summary>
        /// Initializes a new instance of the <see cref="SearchService"/> class.
        /// </summary>
        /// <param name="searchServiceName">IConfiguration provided by DI.</param>
        /// <param name="searchServiceAdminApiKey">TelemetryClient provided by DI.</param>
        /// <param name="searchIndexingIntervalInMinutes">Duration gap for which indexer will run repeatedly.</param>
        /// <param name="storageConnectionString">Azure Table storage connection string.</param>
        /// <param name="searchServiceQueryApiKey">API key required to query search service.</param>
        /// <param name="telemetryClient">Telemetry client for logging events and errors.</param>
        public SearchService(string searchServiceName, string searchServiceAdminApiKey, string searchIndexingIntervalInMinutes, string storageConnectionString, string searchServiceQueryApiKey, TelemetryClient telemetryClient)
        {
            this.telemetryClient = telemetryClient;
            this.searchServiceClient = new SearchServiceClient(searchServiceName, new SearchCredentials(searchServiceAdminApiKey));
            this.storageConnectionString = storageConnectionString;

            if (searchServiceQueryApiKey != null)
            {
                this.searchIndexClient = new SearchIndexClient(searchServiceName, IndexName, new SearchCredentials(searchServiceQueryApiKey));
            }

            this.searchIndexingIntervalInMinutes = Convert.ToInt32(searchIndexingIntervalInMinutes, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Create index, indexer and data source if doesn't exists.
        /// </summary>
        /// <returns>Tracking task.</returns>
        public async Task InitializeAsync()
        {
            try
            {
                await this.CreateIndexAsync().ConfigureAwait(false);
                await this.CreateDataSourceAsync(this.storageConnectionString).ConfigureAwait(false);
                await this.CreateOrRunIndexerAsync().ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"Failed to initialize Azure search service: {ex.Message}", ApplicationInsights.DataContracts.SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Search room or building by name.
        /// </summary>
        /// <param name="searchQuery">Search string.</param>
        /// <returns>List of rooms.</returns>
        public async Task<IList<MeetingRoomEntity>> SearchRoomsAsync(string searchQuery)
        {
            try
            {
                IList<MeetingRoomEntity> rooms = new List<MeetingRoomEntity>();

                SearchParameters searchParam = new SearchParameters
                {
                    OrderBy = new[] { "search.score() desc" },
                    Top = DefaultSearchResultCount,
                };

                var documentSearchResult = await this.searchIndexClient.Documents.SearchAsync<MeetingRoomEntity>(searchQuery + "*").ConfigureAwait(false);
                if (documentSearchResult != null)
                {
                    foreach (SearchResult<MeetingRoomEntity> searchResult in documentSearchResult.Results)
                    {
                        rooms.Add(searchResult.Document);
                    }
                }

                return rooms;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                return null;
            }
        }

        /// <summary>
        /// Create index in Azure search service if it doesn't exists.
        /// </summary>
        /// <returns><see cref="Task"/> that represents index is created if it is not created.</returns>
        private async Task CreateIndexAsync()
        {
            if (await this.searchServiceClient.Indexes.ExistsAsync(IndexName))
            {
                await this.searchServiceClient.Indexes.DeleteAsync(IndexName);
            }

            var tableIndex = new Index()
            {
                Name = IndexName,
                Fields = FieldBuilder.BuildForType<MeetingRoomEntity>(),
            };

            await this.searchServiceClient.Indexes.CreateAsync(tableIndex).ConfigureAwait(false);
        }

        /// <summary>
        /// Add data source if it doesn't exists in Azure search service.
        /// </summary>
        /// <param name="connectionString">connection string to storage.</param>
        /// <returns><see cref="Task"/> that represents data source which is added to Azure search service.</returns>
        private async Task CreateDataSourceAsync(string connectionString)
        {
            if (!await this.searchServiceClient.DataSources.ExistsAsync(DataSourceName))
            {
                var dataSource = DataSource.AzureTableStorage(DataSourceName, connectionString, RoomCollectionStorageProvider.TableName);
                await this.searchServiceClient.DataSources.CreateAsync(dataSource).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Create indexer if it doesn't exists in Azure search service and run if it already exists.
        /// </summary>
        /// <returns>Task that represents indexer is created if not available in Azure search service.</returns>
        private async Task CreateOrRunIndexerAsync()
        {
            if (await this.searchServiceClient.Indexers.ExistsAsync(IndexerName))
            {
                await this.searchServiceClient.Indexers.RunAsync(IndexerName);
            }
            else
            {
                var indexer = new Indexer()
                {
                    Name = IndexerName,
                    DataSourceName = DataSourceName,
                    TargetIndexName = IndexName,
                };

                await this.searchServiceClient.Indexers.CreateAsync(indexer).ConfigureAwait(false);
            }
        }
    }
}
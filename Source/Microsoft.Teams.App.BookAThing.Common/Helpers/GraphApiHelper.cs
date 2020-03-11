// <copyright file="GraphApiHelper.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Helpers
{
    using System.Collections.Generic;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;

    /// <summary>
    /// Methods to perform Graph API calls for GET, POST requests.
    /// </summary>
    public class GraphApiHelper : IGraphApiHelper
    {
        /// <summary>
        /// A factory abstraction for a component that can create HttpClient instances with custom configuration for a given logical name.
        /// </summary>
        private readonly IHttpClientFactory clientFactory;

        /// <summary>
        /// Telemetry client for logging events and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphApiHelper"/> class.
        /// </summary>
        /// <param name="clientFactory">A factory abstraction for a component that can create HttpClient instances with custom configuration for a given logical name.</param>
        /// <param name = "telemetryClient" > Telemetry client for logging events and errors.</param>
        public GraphApiHelper(IHttpClientFactory clientFactory, TelemetryClient telemetryClient)
        {
            this.clientFactory = clientFactory;
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Method to perform HTTP GET requests in Microsoft Graph APIs.
        /// </summary>
        /// <typeparam name="T">Generic type class.</typeparam>
        /// <param name="url">URL to append on base URL for GET.(Example /api/messages).</param>
        /// <param name="token">Authentication token.</param>
        /// <param name="headers">Header parameters.</param>
        /// <returns>API response instance for GET request.</returns>
        public async Task<HttpResponseMessage> GetAsync(string url, string token, Dictionary<string, string> headers = null)
        {
            using (var client = this.clientFactory.CreateClient("GraphApiHelper"))
            {
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                return await client.SendAsync(request).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Method to perform HTTP POST requests in Microsoft Graph APIs.
        /// </summary>
        /// <typeparam name="T">Generic Type class.</typeparam>
        /// <param name="url">URL to append on base URL for POST.(Example /api/messages).</param>
        /// <param name="token">Authentication token.</param>
        /// <param name="payload">request payload in JSON format.</param>
        /// <param name="headers">Header parameters.</param>
        /// <returns>API response instance for POST request.</returns>
        public async Task<HttpResponseMessage> PostAsync(string url, string token, string payload = "", Dictionary<string, string> headers = null)
        {
            using (var client = this.clientFactory.CreateClient("GraphApiHelper"))
            {
                var request = new HttpRequestMessage(HttpMethod.Post, url);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                if (!string.IsNullOrEmpty(payload))
                {
                    request.Content = new StringContent(payload, Encoding.UTF8, "application/json");
                }

                return await client.SendAsync(request).ConfigureAwait(false);
            }
        }
    }
}
// <copyright file="UserConfigurationProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Common.Providers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Teams.App.BookAThing.Common.Models.Error;
    using Microsoft.Teams.App.BookAThing.Common.Models.Response;
    using Microsoft.Teams.Apps.BookAThing.Common.Helpers;
    using Newtonsoft.Json;

    /// <summary>
    /// Exposes methods for fetching user specific data.
    /// </summary>
    public class UserConfigurationProvider : IUserConfigurationProvider
    {
        /// <summary>
        /// API helper service for making post and get calls to Graph.
        /// </summary>
        private readonly IGraphApiHelper apiHelper;

        /// <summary>
        /// Telemetry client for logging events and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Graph API to get supported time zones configured in user mailbox.
        /// </summary>
        private readonly string supportedTimeZoneGraphEndpointUrl = "/v1.0/me/outlook/supportedTimeZones(TimeZoneStandard=microsoft.graph.timeZoneStandard'Iana')";

        /// <summary>
        /// Initializes a new instance of the <see cref="UserConfigurationProvider"/> class.
        /// </summary>
        /// <param name="apiHelper">Api helper service for making post and get calls to Microsoft Graph APIs.</param>
        /// <param name="telemetryClient">Telemetry client for logging events and errors.</param>
        public UserConfigurationProvider(IGraphApiHelper apiHelper, TelemetryClient telemetryClient)
        {
            this.apiHelper = apiHelper;
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Get supported time zones for signed in user.
        /// </summary>
        /// <param name="token">Active Directory access token.</param>
        /// <returns>List of supported time zones.</returns>
        public async Task<SupportedTimeZoneResponse> GetSupportedTimeZoneAsync(string token)
        {
            try
            {
                var httpResponseMessage = await this.apiHelper.GetAsync(this.supportedTimeZoneGraphEndpointUrl, token).ConfigureAwait(false);
                var content = await httpResponseMessage.Content.ReadAsStringAsync();

                if (httpResponseMessage.IsSuccessStatusCode)
                {
                    return JsonConvert.DeserializeObject<SupportedTimeZoneResponse>(content);
                }
                else
                {
                    var errorResponse = JsonConvert.DeserializeObject<ErrorResponse>(content);

                    this.telemetryClient.TrackTrace($"Graph API failure- url: {this.supportedTimeZoneGraphEndpointUrl}, response-code: {errorResponse.Error.StatusCode}, response-content: {errorResponse.Error.ErrorMessage}, request-id: {errorResponse.Error.InnerError.RequestId}", SeverityLevel.Warning);
                    var failureResponse = new SupportedTimeZoneResponse
                    {
                        StatusCode = httpResponseMessage.StatusCode,
                        ErrorResponse = errorResponse,
                    };

                    return failureResponse;
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }
    }
}

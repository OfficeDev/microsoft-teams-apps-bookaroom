// <copyright file="UserData.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Models
{
    /// <summary>
    /// Class containing properties which are user specific.
    /// </summary>
    public class UserData
    {
        /// <summary>
        /// Gets or sets a value indicating whether welcome card sent.
        /// </summary>
        public bool? IsWelcomeCardSent { get; set; }
    }
}

// <copyright file="JwtClaim.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Models
{
    /// <summary>
    /// Claims which are added in JWT token.
    /// </summary>
    public class JwtClaim
    {
        /// <summary>
        /// Gets or sets Active Directory object Id of user.
        /// </summary>
        public string UserObjectIdentifer { get; set; }

        /// <summary>
        /// Gets or sets activity Id.
        /// </summary>
        public string FromId { get; set; }

        /// <summary>
        /// Gets or sets service url of bot.
        /// </summary>
        public string ServiceUrl { get; set; }
    }
}

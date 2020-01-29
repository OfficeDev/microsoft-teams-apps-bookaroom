// <copyright file="ActivityEntity.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Models.TableEntities
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Activity table entity class for storing bot activity Id for card updation.
    /// </summary>
    public class ActivityEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets user Active Directory object Id.
        /// </summary>
        public string UserAdObjectId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        /// Gets or sets activity reference Id generated before sending card to user.
        /// </summary>
        public string ActivityReferenceId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets activity Id.
        /// </summary>
        public string ActivityId { get; set; }
    }
}

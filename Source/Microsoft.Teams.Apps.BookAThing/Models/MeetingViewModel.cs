// <copyright file="MeetingViewModel.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Models
{
    using Microsoft.Teams.Apps.BookAThing.Common.Models.TableEntities;

    /// <summary>
    /// Class containing properties to be parsed from activity value.
    /// </summary>
    public class MeetingViewModel : UserFavoriteRoomEntity
    {
        /// <summary>
        /// Gets or sets bot command text.
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether room is favorite.
        /// </summary>
        public bool IsFavourite { get; set; }

        /// <summary>
        /// Gets or sets reply to activity Id.
        /// </summary>
        public string ReplyTo { get; set; }

        /// <summary>
        /// Gets or sets start date of meeting.
        /// </summary>
        public string StartDateTime { get; set; }

        /// <summary>
        /// Gets or sets end date of meeting.
        /// </summary>
        public string EndDateTime { get; set; }

        /// <summary>
        /// Gets or sets Active Directory object Id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets web link of meeting.
        /// </summary>
        public string WebLink { get; set; }

        /// <summary>
        /// Gets or sets meeting Id.
        /// </summary>
        public string MeetingId { get; set; }

        /// <summary>
        /// Gets or sets subject of meeting.
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Gets or sets status of room.
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// Gets or sets unique GUID to recognize previous activity which needs to be updated.
        /// </summary>
        public string ActivityReferenceId { get; set; }
    }
}

// <copyright file="BotCommands.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing
{
    /// <summary>
    /// Bot commands.
    /// </summary>
    public static class BotCommands
    {
        /// <summary>
        /// Book room bot command which will show user favorites list card.
        /// </summary>
        public const string BookAMeeting = "BOOK ROOM";

        /// <summary>
        /// Create meeting command.
        /// </summary>
        public const string CreateMeeting = "CREATE MEETING";

        /// <summary>
        /// Cancel meeting command.
        /// </summary>
        public const string CancelMeeting = "CANCEL MEETING";

        /// <summary>
        /// Manage favorites command.
        /// </summary>
        public const string ManageFavorites = "MANAGE FAVORITES";

        /// <summary>
        /// Add favorite command.
        /// </summary>
        public const string AddFavorite = "ADD FAVORITES";

        /// <summary>
        /// Login command.
        /// </summary>
        public const string Login = "SIGN IN";

        /// <summary>
        /// Logout command.
        /// </summary>
        public const string Logout = "SIGN OUT";

        /// <summary>
        /// Help command.
        /// </summary>
        public const string Help = "HELP";

        /// <summary>
        /// Show favorite task module command.
        /// </summary>
        public const string ShowFavoriteTaskModule = "SHOW FAVORITE TASK MODULE";

        /// <summary>
        /// Show other rooms task module command.
        /// </summary>
        public const string ShowOtherRoomsTaskModule = "SHOW OTHER ROOMS TASK MODULE";

        /// <summary>
        /// Refresh favorite list command.
        /// </summary>
        public const string RefreshList = "REFRESH FAVORITE LIST";

        /// <summary>
        /// Command from task module after successful meeting creation.
        /// </summary>
        public const string MeetingFromTaskModule = "MEETING FROM TASK MODULE";
    }
}

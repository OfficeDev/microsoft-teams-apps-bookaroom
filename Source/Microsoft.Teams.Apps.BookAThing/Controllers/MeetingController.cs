// <copyright file="MeetingController.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Controllers
{
    using Microsoft.AspNetCore.Mvc;

    /// <summary>
    /// Controller class for add favorites and other room views.
    /// </summary>
    [Route("[controller]/[action]")]
    public class MeetingController : Controller
    {
        /// <summary>
        /// Add favorites view for managing favorite rooms of user.
        /// </summary>
        /// <returns> Client App View. </returns>
        public ActionResult AddFavourite()
        {
            return this.View();
        }

        /// <summary>
        /// Other room view for booking non-favorite room for specific time interval.
        /// </summary>
        /// <returns> Other Room View. </returns>
        public ActionResult OtherRoom()
        {
            return this.View();
        }
    }
}
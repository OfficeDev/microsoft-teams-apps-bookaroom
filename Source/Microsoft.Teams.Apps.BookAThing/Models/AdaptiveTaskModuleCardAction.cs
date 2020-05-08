// <copyright file="AdaptiveTaskModuleCardAction.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Defines model for opening task module.
    /// </summary>
    public class AdaptiveTaskModuleCardAction
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AdaptiveTaskModuleCardAction"/> class.
        /// </summary>
        /// <param name="skillId"> Microsoft app id to embedded in card actions. </param>
        public AdaptiveTaskModuleCardAction(string skillId)
        {
            this.SkillId = skillId;
        }

        /// <summary>
        /// Gets or sets action type for button.
        /// </summary>
        [JsonProperty("type")]
        public string Type
        {
            get
            {
                return "task/fetch";
            }
            set => this.Type = "task/fetch";
        }

        /// <summary>
        /// Gets or sets bot command to be used by bot for processing user inputs.
        /// </summary>
        [JsonProperty("text")]
        public string Text { get; set; }

        /// <summary>
        /// Gets or sets unique GUID related to activity Id from ActivityEntities table.
        /// </summary>
        [JsonProperty("activityReferenceId")]
        public string ActivityReferenceId { get; set; }

        /// <summary>
        /// Gets or sets skillId.
        /// </summary>
        [JsonProperty("skillId")]
        public string SkillId { get; set; }
    }
}

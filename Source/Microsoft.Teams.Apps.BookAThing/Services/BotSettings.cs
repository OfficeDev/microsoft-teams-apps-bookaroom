// <copyright file="BotSettings.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Services
{
    using System.Collections.Generic;

    public class BotSettings
    {
        public Dictionary<string, LanguageModel> LanguageModels { get; set; } = new Dictionary<string, LanguageModel>();

        public string DefaultLocale { get; set; }
    }
}

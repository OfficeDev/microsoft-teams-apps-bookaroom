// <copyright file="BotServices.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Services
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using Microsoft.Bot.Builder.AI.Luis;

    public class BotServices
    {
        public BotServices(BotSettings settings)
        {
            foreach (var pair in settings.LanguageModels)
            {
                var language = pair.Key;
                var languageModel = pair.Value;

                if (languageModel != null)
                {
                    var luisApp = new LuisApplication(languageModel.LuisAppId, languageModel.LuisApiKey, languageModel.LuisApiHost);
                    LanguageModels.Add(language, new LuisRecognizer(luisApp));
                }
            }
        }

        public Dictionary<string, LuisRecognizer> LanguageModels { get; set; } = new Dictionary<string, LuisRecognizer>();

        public LuisRecognizer GetLanguageModels()
        {
            // Get cognitive models for locale
            var locale = CultureInfo.CurrentUICulture.Name.ToLower();

            var languageModel = this.LanguageModels.ContainsKey(locale)
                ? this.LanguageModels[locale]
                : this.LanguageModels.Where(key => key.Key.StartsWith(locale.Substring(0, 2))).FirstOrDefault().Value
                ?? throw new Exception($"There's no matching locale for '{locale}' or its root language '{locale.Substring(0, 2)}'. " +
                                        "Please review your available locales in your cognitivemodels.json file.");

            return languageModel;
        }
    }
}
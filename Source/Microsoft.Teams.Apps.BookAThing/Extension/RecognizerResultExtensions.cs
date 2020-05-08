// <copyright file="RecognizerResultExtensions.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.BookAThing.Extension
{
    using Microsoft.Bot.Builder;

    /// <summary>
    /// Extension methods for RecognizerResult.
    /// </summary>
    public static class RecognizerResultExtensions
    {
        /// <summary>
        /// Returns top intent and corresponding score.
        /// </summary>
        /// <param name="result">RecognizerResult.</param>
        /// <returns>Tuple containing top intent and corresponding score.</returns>
        public static (string, double) GetTopIntentAndScore(this RecognizerResult result)
        {
            string topIntent = "None";
            var topIntentScore = 0.0;
            foreach (var entry in result.Intents)
            {
                if (entry.Value.Score > topIntentScore)
                {
                    topIntent = entry.Key;
                    topIntentScore = entry.Value.Score.Value;
                }
            }

            return (topIntent, topIntentScore);
        }
    }
}

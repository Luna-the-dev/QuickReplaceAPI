/******************************************************************
 * Copyright (c) QuickReplace, LLC. All Rights Reserved.
 * 
 * This file is part of the QuickReplace Developer Library. You may not
 * use this file for commercial or business use except in compliance with
 * the QuickReplace license. You may obtain a copy of the license at:
 * 
 *   https://www.quickreplace.io/pricing
 * 
 * You may view the terms of this license at:
 * 
 *   https://www.quickreplace.io/eula
 * 
 * This file or any others in the QuickReplace Developer Library may not
 * be modified, shared, or distributed outside of the terms of the license
 * agreement by organizations or individuals who are covered by a license.
 * 
 *****************************************************************/

using TextReplaceAPI.DataTypes;

namespace TextReplaceAPI.Core.AhoCorasick
{
    internal static class AhoCorasickMatcher
    {
        public static AhoCorasickStringSearcher CreateMatcher(Dictionary<string, string> replacePhrases, bool caseSensitive)
        {
            // construct the automaton and fill it with the phrases to search for
            // also create a list of the replacement phrases to go alongside the 
            var matcher = new AhoCorasickStringSearcher(caseSensitive);
            foreach (var searchWord in replacePhrases)
            {
                matcher.AddItem(searchWord.Key);
            }
            matcher.CreateFailureFunction();

            return matcher;
        }
    }
}

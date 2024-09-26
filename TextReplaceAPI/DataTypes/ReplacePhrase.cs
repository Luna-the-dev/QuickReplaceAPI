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

using CsvHelper.Configuration.Attributes;

namespace TextReplaceAPI.DataTypes
{
    /// <summary>
    /// Wrapper class for the replace phrases dictionary.
    /// This wrapper exists only to read in data with CsvHelper.
    /// Note: Keep the variables in the constructor the exact same as the fields.
    /// This is to make CsvHelper work without a default constructor.
    /// </summary>
    internal class ReplacePhrase
    {
        [Index(0)]
        public string Item1 { get; set; }

        [Index(1)]
        public string Item2 { get; set; }

        public ReplacePhrase()
        {
            Item1 = string.Empty;
            Item2 = string.Empty;
        }

        public ReplacePhrase(string item1, string item2)
        {
            Item1 = item1;
            Item2 = item2;
        }
    }
}

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

namespace TextReplaceAPI.Core.Validation
{
    internal class DataValidation
    {
        private const string INVALID_DELIMITER_CHARS = "\n";
        private const string INVALID_SUFFIX_CHARS = "<>:\"/\\|?*\n\t";

        /// <summary>
        /// Checks if the delimiter string is empty or contains any invalid characters.
        /// </summary>
        /// <param name="delimiter"></param>
        /// <returns>True if the string is empty or does not contain any invalid characters.</returns>
        public static bool IsDelimiterValid(string delimiter)
        {
            if (delimiter == string.Empty)
            {
                return false;
            }

            foreach (char c in delimiter)
            {
                if (INVALID_DELIMITER_CHARS.Contains(c))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Checks if the suffix string contains any invalid characters.
        /// </summary>
        /// <param name="suffix"></param>
        /// <returns>True if the string is empty or does not contain any invalid characters.</returns>
        public static bool IsSuffixValid(string suffix)
        {
            foreach (char c in suffix)
            {
                if (INVALID_SUFFIX_CHARS.Contains(c) || char.IsControl(c))
                {
                    return false;
                }
            }
            return true;
        }
    }
 }

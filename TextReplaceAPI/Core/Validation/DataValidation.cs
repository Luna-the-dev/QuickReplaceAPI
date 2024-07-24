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

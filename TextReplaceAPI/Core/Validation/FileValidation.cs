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

using System.Diagnostics;

namespace TextReplaceAPI.Core.Validation
{
    internal class FileValidation
    {
        /// <summary>
        /// Checks to see if the provided file exists and is readable.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>True if it is both readable and exists</returns>
        public static bool IsInputFileReadable(string fileName)
        {
            try
            {
                File.Open(fileName, FileMode.Open, FileAccess.Read).Dispose();
                return true;
            }
            catch (ArgumentException)
            {
                Debug.WriteLine("File name is empty.");
                return false;
            }
            catch
            {
                Debug.WriteLine("File could not be opened.");
                return false;
            }
        }

        /// <summary>
        /// Checks to see if the provided file exists and is readable/writable.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>True if it is both readable and exists</returns>
        public static bool IsInputFileReadWriteable(string fileName)
        {
            try
            {
                File.Open(fileName, FileMode.Open, FileAccess.ReadWrite).Dispose();
                return true;
            }
            catch (ArgumentException)
            {
                Debug.WriteLine("File name is empty.");
                return false;
            }
            catch
            {
                Debug.WriteLine("File could not be opened.");
                return false;
            }
        }

        /// <summary>
        /// Checks to see if the replace file is of a supported type
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>False if the file type is not supported</returns>
        public static bool IsReplaceFileTypeValid(string fileName)
        {
            string extension = Path.GetExtension(fileName).ToLower();
            return extension switch
            {
                ".csv" or ".tsv" or ".xlsx" or ".txt" or ".text" => true,
                _ => false
            };
        }

        /// <summary>
        /// Checks to see if the source file is of a supported type
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>False if the file type is not supported</returns>
        public static bool IsSourceFileTypeValid(string fileName)
        {
            if (Path.GetExtension(fileName).Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase) ||
                Path.GetExtension(fileName).Equals(".docx", StringComparison.CurrentCultureIgnoreCase))
            {
                return true;
            }

            if (IsFileNonBinary(fileName))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Checks to see if the file conversion from source to output is valid
        /// </summary>
        /// <param name="sourceFile"></param>
        /// <param name="outputFile"></param>
        /// <returns>Returns false if the file conversion is not valid</returns>
        public static bool IsFileConversionValid(string sourceFile, string outputFile)
        {
            var sourceExtension = Path.GetExtension(sourceFile);
            var outputExtension = Path.GetExtension(outputFile);

            // source is excel, output is not
            if (Path.GetExtension(sourceFile).Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase) &&
                !Path.GetExtension(outputFile).Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase))
            {
                return false;
            }

            // output is excel, source is not
            if (!Path.GetExtension(sourceFile).Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase) &&
                Path.GetExtension(outputFile).Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase))
            {
                return false;
            }

            return true;
        }

        public static bool IsTextFile(string fileName)
        {
            if (Path.GetExtension(fileName).Equals(".txt", StringComparison.CurrentCultureIgnoreCase) ||
                Path.GetExtension(fileName).Equals(".text", StringComparison.CurrentCultureIgnoreCase))
            {
                return true;
            }
            return false;
        }

        public static bool IsExcelFile(string fileName)
        {
            if (Path.GetExtension(fileName).Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase))
            {
                return true;
            }
            return false;
        }

        public static bool IsDocxFile(string fileName)
        {
            if (Path.GetExtension(fileName).Equals(".docx", StringComparison.CurrentCultureIgnoreCase))
            {
                return true;
            }
            return false;
        }

        public static bool IsFileNonBinary(string fileName)
        {
            return !IsFileBinary(fileName);
        }

        public static bool IsFileBinary(string fileName)
        {
            const int charsToCheck = 4000;
            const char nulChar = '\0';

            using var streamReader = new StreamReader(fileName);

            for (var i = 0; i < charsToCheck; i++)
            {
                if (streamReader.EndOfStream)
                {
                    return false;
                }

                var ch = (char)streamReader.Read();
                if (ch == nulChar)
                {
                    return true;
                }
            }

            return false;
        }
    }
}

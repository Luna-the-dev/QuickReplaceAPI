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

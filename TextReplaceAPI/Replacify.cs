using TextReplaceAPI.Core.Data;
using TextReplaceAPI.Core.Validation;
using TextReplaceAPI.MVVM.Model;

namespace TextReplaceAPI
{
    public class Replacify
    {
        // key is the phrase to replace, value is what it is being replaced with
        private Dictionary<string, string> _replacePhrases = [];
        public Dictionary<string, string> ReplacePhrases
        {
            get { return _replacePhrases; }
            set { _replacePhrases = value; }
        }

        private IEnumerable<SourceFile> _sourceFiles;
        public IEnumerable<SourceFile> SourceFiles
        {
            get { return _sourceFiles; }
            set { _sourceFiles = value; }
        }

        /// <summary>
        /// Initializes the Replacify class.
        /// </summary>
        /// <param name="replacementsFileName"></param>
        /// <exception cref="IOException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="NotSupportedException"></exception>
        /// <exception cref="CsvHelper.MissingFieldException"></exception>
        public Replacify(
            string replacementsFileName,
            IEnumerable<string> sourceFileNames,
            IEnumerable<string> outputFileNames)
		{
            _replacePhrases = ParseReplacements(replacementsFileName);

            if (AreSourceFileTypesValid(sourceFileNames) == false)
            {
                throw new NotSupportedException("Invalid source file: the only supported file types are .csv, .tsv, .xlsx., .txt, and .text");
            }

            if (AreOutputFileTypesValid(sourceFileNames) == false)
            {
                throw new NotSupportedException("Invalid output file: the only supported file types are .csv, .tsv, .xlsx., .txt, and .text");
            }

            // zip the source file names and the output file names together and then combine the names into SourceFile objects
            _sourceFiles = ZipSourceFiles(sourceFileNames, outputFileNames);
        }

        /// <summary>
        /// Searches through a list of source files, looking for instances of keys from 
        /// the ReplacePhrases dict, replacing them with the associated value, and then
        /// saving the resulting text off to a list of destination files.
        /// 
        /// Note: if the "throwExceptions" flag is set to false and writing to one of the files failed,
        /// the method will continue to write to the other files in the list.
        /// </summary>
        /// <param name="wholeWord"></param>
        /// <param name="caseSensitive"></param>
        /// <param name="preserveCase"></param>
        /// <param name="styling"></param>
        /// <param name="throwExceptions"></param>
        /// <returns></returns>
        public bool PerformReplacements(
            bool wholeWord,
            bool caseSensitive,
            bool preserveCase,
            OutputFileStyling? styling = null,
            bool throwExceptions = true)
        {
            if (throwExceptions)
            {
                // will throw exception if something goes wrong
                OutputData.PerformReplacementsThrowExceptions(
                    ReplacePhrases, SourceFiles, wholeWord, caseSensitive, preserveCase, styling);
                return true;
            }

            // catches any exceptions if something goes wrong and continues to write to
            // the remaining files. will only throw an ArgumentException if SourceFiles is empty.
            // returns false if something went wrong, true if all files wrote successfully.
            return OutputData.PerformReplacements(
                ReplacePhrases, SourceFiles, wholeWord, caseSensitive, preserveCase, styling);
        }

        /// <summary>
        /// Searches through a list of source files, looking for instances of keys from 
        /// the ReplacePhrases dict, replacing them with the associated value, and then
        /// saving the resulting text off to a list of destination files.
        /// 
        /// Note: if the "throwExceptions" flag is set to false and writing to one of the files failed,
        /// the method will continue to write to the other files in the list.
        /// </summary>
        /// <param name="wholeWord"></param>
        /// <param name="caseSensitive"></param>
        /// <param name="preserveCase"></param>
        /// <param name="styling"></param>
        /// <param name="throwExceptions"></param>
        /// <returns>True is all replacements were made successfully, false if something went wrong.</returns>
        public static bool PerformReplacements(
            Dictionary<string, string> replacePhrases,
            IEnumerable<SourceFile> sourceFiles,
            bool wholeWord,
            bool caseSensitive,
            bool preserveCase,
            OutputFileStyling? styling = null,
            bool throwExceptions = true)
        {
            if (throwExceptions)
            {
                // will throw exception if something goes wrong
                OutputData.PerformReplacementsThrowExceptions(
                    replacePhrases, sourceFiles, wholeWord, caseSensitive, preserveCase, styling);
                return true;
            }

            // catches any exceptions if something goes wrong and continues to write to
            // the remaining files. will only throw an ArgumentException if SourceFiles is empty.
            // returns false if something went wrong, true if all files wrote successfully.
            return OutputData.PerformReplacements(
                replacePhrases, sourceFiles, wholeWord, caseSensitive, preserveCase, styling);
        }

        /// <summary>
        /// Parses the replacements from a file. Supports excel files as well as value-seperated files
        /// Supported file types are: .csv, .tsv, .xlsx., .txt, and .text
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>
        /// A dictionary of pairs of the values from the file. If one of the lines in the file has an
        /// incorrect number of values or if the operation fails for another reason, return an empty list.
        /// </returns>
        /// <exception cref="NotSupportedException"></exception>
        public static Dictionary<string, string> ParseReplacements(string fileName)
        {
            if (FileValidation.IsReplaceFileTypeValid(fileName) == false)
            {
                throw new NotSupportedException($"File type {Path.GetExtension(fileName).ToLower()} is not supported as a replacements file.");
            }

            return ReplacementsHelper.ParseReplacements(fileName);
        }

        /// <summary>
        /// Zips an IEnumberable of source file names and an IEnumerable of output file names together
        /// an IEnumerable of SourceFile objects
        /// </summary>
        /// <param name="sourceFileNames"></param>
        /// <param name="outputFileNames"></param>
        /// <returns>
        /// An IEnumberable<SourceFile> containing the source file names,
        /// output file names, and -1 for the number of replacements made.
        /// </returns>
        /// <exception cref="ArgumentException">The number of source files does not equal the number of output files.</exception>
        public static IEnumerable<SourceFile> ZipSourceFiles(IEnumerable<string> sourceFileNames, IEnumerable<string> outputFileNames)
        {
            if (sourceFileNames.Count() != outputFileNames.Count())
            {
                throw new ArgumentException("The number of source files does not equal the number of output files.");
            }

            // zip the source file names and the output file names together and then combine the names into SourceFile objects
            return sourceFileNames
                .Zip(outputFileNames, (s, o) => new { SourceFileName = s, OutputFIleName = o })
                .Select(x => new SourceFile(x.SourceFileName, x.OutputFIleName, -1));
        }

        /// <summary>
        /// Determines whether a replacement file name ends with a valid file type.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>True if the file ends in a valid file type.</returns>
        public static bool IsReplacementFileTypeValid(string fileName)
        {
            return FileValidation.IsReplaceFileTypeValid(fileName);
        }

        /// <summary>
        /// Determines whether an IEnumerable of source file names all end with valid file types.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>True if all files end in valid file types.</returns>
        public static bool AreSourceFileTypesValid(IEnumerable<string> fileNames)
        {
            foreach(var fileName in fileNames)
            {
                if (IsSourceFileTypeValid(fileName) == false)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Determines whether a source file name ends with a valid file type.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>True if the file ends in a valid file type.</returns>
        public static bool IsSourceFileTypeValid(string fileName)
        {
            return FileValidation.IsSourceFileTypeValid(fileName);
        }

        /// <summary>
        /// Determines whether an IEnumerable of source file names all end with valid file types.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>True if all files end in valid file types.</returns>
        public static bool AreOutputFileTypesValid(IEnumerable<string> fileNames)
        {
            foreach (var fileName in fileNames)
            {
                if (IsOutputFileTypeValid(fileName) == false)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Determines whether an output file name ends with a valid file type.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>True if the file ends in a valid file type.</returns>
        public static bool IsOutputFileTypeValid(string fileName)
        {
            return FileValidation.IsOutputFileTypeValid(fileName);
        }
    }
}

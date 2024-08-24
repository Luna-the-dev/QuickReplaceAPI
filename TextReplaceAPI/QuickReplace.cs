﻿using TextReplaceAPI.Core.AhoCorasick;
using TextReplaceAPI.Core.Helpers;
using TextReplaceAPI.Core.Validation;
using TextReplaceAPI.DataTypes;
using TextReplaceAPI.Exceptions;

namespace TextReplaceAPI
{
    public class QuickReplace
    {
        // key is the phrase to replace, value is what it is being replaced with
        private Dictionary<string, string> _replacePhrases = [];
        public Dictionary<string, string> ReplacePhrases
        {
            get { return _replacePhrases; }
            set { _replacePhrases = value; }
        }

        private List<SourceFile> _sourceFiles;
        public List<SourceFile> SourceFiles
        {
            get { return _sourceFiles; }
            set { _sourceFiles = value; }
        }

        private AhoCorasickStringSearcher? _matcher = null;

        /// <summary>
        /// Initializes a Dictionary<string, string> of replace phrases and an
        /// IEnumberable<SourceFile> containing the source and output file names.
        /// </summary>
        /// <param name="replacementsFileName"></param>
        /// <param name="sourceFileNames"></param>
        /// <param name="outputFileNames"></param>
        /// <param name="validateSrcFiles"></param>
        /// <exception cref="InvalidFileTypeException">
        /// A file type is not supported. See documentation for a list of supported file types.
        /// </exception>
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="PathTooLongException"></exception>
        /// <exception cref="DirectoryNotFoundException"></exception>
        /// <exception cref="FileNotFoundException"></exception>
        /// <exception cref="IOException"></exception>
        /// <exception cref="UnauthorizedAccessException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public QuickReplace(
            string replacementsFileName,
            IEnumerable<string> sourceFileNames,
            IEnumerable<string> outputFileNames,
            bool validateSrcFiles = false)
		{
            _replacePhrases = ParseReplacements(replacementsFileName);

            if (validateSrcFiles)
            {
                if (AreFilesValid(sourceFileNames, outputFileNames) == false)
                {
                    throw new InvalidFileTypeException("Invalid source file: the only supported file types are .csv, .tsv, .xlsx., .txt, and .text");
                }
            }

            // zip the source file names and the output file names together and then combine the names into SourceFile objects
            _sourceFiles = ZipSourceFiles(sourceFileNames, outputFileNames);
        }

        /// <summary>
        /// Initializes a Dictionary<string, string> of replace phrases and an
        /// IEnumberable<SourceFile> containing the source and output file names.
        /// </summary>
        /// <param name="replacements"></param>
        /// <param name="sourceFileNames"></param>
        /// <param name="outputFileNames"></param>
        /// <param name="validateSrcFiles"></param>
        /// <exception cref="InvalidFileTypeException">
        /// A file type is not supported. See documentation for a list of supported file types.
        /// </exception>
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="PathTooLongException"></exception>
        /// <exception cref="DirectoryNotFoundException"></exception>
        /// <exception cref="FileNotFoundException"></exception>
        /// <exception cref="IOException"></exception>
        /// <exception cref="UnauthorizedAccessException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public QuickReplace(
            Dictionary<string, string> replacements,
            IEnumerable<string> sourceFileNames,
            IEnumerable<string> outputFileNames,
            bool validateSrcFiles = false)
        {
            _replacePhrases = replacements;

            if (validateSrcFiles)
            {
                if (AreFilesValid(sourceFileNames, outputFileNames) == false)
                {
                    throw new InvalidFileTypeException("Invalid source file: the only supported file types are .csv, .tsv, .xlsx., .txt, and .text");
                }
            }

            // zip the source file names and the output file names together and then combine the names into SourceFile objects
            _sourceFiles = ZipSourceFiles(sourceFileNames, outputFileNames);
        }

        /// <summary>
        /// Initializes a Dictionary<string, string> of replace phrases, an IEnumberable<SourceFile>
        /// containing the source and output file names, and the Aho-Corasick matcher if the
        /// preGenerateMatcher argument is true.
        /// 
        /// Only use this contructor if you would like to pre-generate the Aho-Corasick matcher.
        /// This front-loads much of the processing time upon instantiation of this class's object
        /// rather than when the Replace method is called.
        /// </summary>
        /// <param name="replacementsFileName"></param>
        /// <param name="sourceFileNames"></param>
        /// <param name="outputFileNames"></param>
        /// <param name="preGenerateMatcher"></param>
        /// <param name="caseSensitive"></param>
        /// <param name="validateSrcFiles"></param>
        /// <exception cref="InvalidFileTypeException">
        /// A file type is not supported. See documentation for a list of supported file types.
        /// </exception>
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="PathTooLongException"></exception>
        /// <exception cref="DirectoryNotFoundException"></exception>
        /// <exception cref="FileNotFoundException"></exception>
        /// <exception cref="IOException"></exception>
        /// <exception cref="UnauthorizedAccessException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public QuickReplace(
            string replacementsFileName,
            IEnumerable<string> sourceFileNames,
            IEnumerable<string> outputFileNames,
            bool preGenerateMatcher,
            bool caseSensitive,
            bool validateSrcFiles = false)
        {
            _replacePhrases = ParseReplacements(replacementsFileName);

            if (validateSrcFiles)
            {
                if (AreFilesValid(sourceFileNames, outputFileNames) == false)
                {
                    throw new InvalidFileTypeException("Invalid source file: the only supported file types are .csv, .tsv, .xlsx., .txt, and .text");
                }
            }

            // zip the source file names and the output file names together and then combine the names into SourceFile objects
            _sourceFiles = ZipSourceFiles(sourceFileNames, outputFileNames);

            // generate the matcher
            if (preGenerateMatcher)
            {
                GenerateMatcher(caseSensitive);
            }
        }

        /// <summary>
        /// Initializes a Dictionary<string, string> of replace phrases, an IEnumberable<SourceFile>
        /// containing the source and output file names, and the Aho-Corasick matcher if the
        /// preGenerateMatcher argument is true.
        /// 
        /// Only use this contructor if you would like to pre-generate the Aho-Corasick matcher.
        /// This front-loads much of the processing time upon instantiation of this class's object
        /// rather than when the Replace method is called.
        /// </summary>
        /// <param name="replacements"></param>
        /// <param name="sourceFileNames"></param>
        /// <param name="outputFileNames"></param>
        /// <param name="preGenerateMatcher"></param>
        /// <param name="caseSensitive"></param>
        /// <param name="validateSrcFiles"></param>
        /// <exception cref="InvalidFileTypeException">
        /// A file type is not supported. See documentation for a list of supported file types.
        /// </exception>
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="PathTooLongException"></exception>
        /// <exception cref="DirectoryNotFoundException"></exception>
        /// <exception cref="FileNotFoundException"></exception>
        /// <exception cref="IOException"></exception>
        /// <exception cref="UnauthorizedAccessException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public QuickReplace(
            Dictionary<string, string> replacements,
            IEnumerable<string> sourceFileNames,
            IEnumerable<string> outputFileNames,
            bool preGenerateMatcher,
            bool caseSensitive,
            bool validateSrcFiles = false)
        {
            _replacePhrases = replacements;

            if (validateSrcFiles)
            {
                if (AreFilesValid(sourceFileNames, outputFileNames) == false)
                {
                    throw new InvalidFileTypeException("Invalid source file: the only supported file types are .csv, .tsv, .xlsx., .txt, and .text");
                }
            }

            // zip the source file names and the output file names together and then combine the names into SourceFile objects
            _sourceFiles = ZipSourceFiles(sourceFileNames, outputFileNames);

            // generate the matcher
            if (preGenerateMatcher)
            {
                GenerateMatcher(caseSensitive);
            }
        }

        /// <summary>
        /// Searches through a list of source files, looking for instances of keys from 
        /// the ReplacePhrases dict, replacing them with the associated value, and then
        /// saving the resulting text off to a list of destination files.
        /// 
        /// If the "throwExceptions" flag is set to false and writing to one of the files failed,
        /// the method will prevent the exception from bubbling up to the caller and continue
        /// to write to the other files in the list.
        /// </summary>
        /// <param name="replacements"></param>
        /// <param name="sourceFiles"></param>
        /// <param name="wholeWord"></param>
        /// <param name="caseSensitive"></param>
        /// <param name="preserveCase"></param>
        /// <param name="styling"></param>
        /// <param name="throwExceptions"></param>
        /// <returns>True is all replacements were made successfully, false if something went wrong.</returns>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="PathTooLongException"></exception>
        /// <exception cref="DirectoryNotFoundException"></exception>
        /// <exception cref="FileNotFoundException"></exception>
        /// <exception cref="IOException"></exception>
        /// <exception cref="UnauthorizedAccessException"></exception>
        /// <exception cref="System.Security.SecurityException"></exception>
        /// <exception cref="InvalidXmlStructureException">
        /// The XML data within a .docx or .xlsx file has an incorrect structure and could not be parsed.
        /// </exception>
        public static bool Replace(
            Dictionary<string, string> replacements,
            IEnumerable<SourceFile> sourceFiles,
            bool wholeWord = false,
            bool caseSensitive = false,
            bool preserveCase = false,
            Styling? styling = null,
            bool throwExceptions = true)
        {
            // If throwExceptions is true, Replace() will allow exceptions to bubble up to the caller.
            // If it is false, Replace() will catch any exceptions and continues to write to
            // the remaining files. It will only allow an ArgumentException to bubble up to the caller if SourceFiles is empty.
            // This returns false if something went wrong, and true if all files wrote successfully.
            return OutputHelper.Replace(
                replacements, sourceFiles, wholeWord, caseSensitive, preserveCase, throwExceptions, styling);
        }

        /// <summary>
        /// Searches through a list of source files, looking for instances of keys from 
        /// the ReplacePhrases dict, replacing them with the associated value, and then
        /// saving the resulting text off to a list of destination files.
        /// 
        /// If the "throwExceptions" flag is set to false and writing to one of the files failed,
        /// the method will continue to write to the other files in the list.
        /// </summary>
        /// <param name="wholeWord"></param>
        /// <param name="caseSensitive"></param>
        /// <param name="preserveCase"></param>
        /// <param name="styling"></param>
        /// <param name="throwExceptions"></param>
        /// <returns>True is all replacements were made successfully, false if something went wrong.</returns>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="PathTooLongException"></exception>
        /// <exception cref="DirectoryNotFoundException"></exception>
        /// <exception cref="FileNotFoundException"></exception>
        /// <exception cref="IOException"></exception>
        /// <exception cref="UnauthorizedAccessException"></exception>
        /// <exception cref="System.Security.SecurityException"></exception>
        /// <exception cref="InvalidXmlStructureException">
        /// The XML data within a .docx or .xlsx file has an incorrect structure and could not be parsed.
        /// </exception>
        public bool Replace(
            bool wholeWord = false,
            bool caseSensitive = false,
            bool preserveCase = false,
            Styling? styling = null,
            bool throwExceptions = true)
        {
            // If throwExceptions is true, Replace() will allow exceptions to bubble up to the caller.
            // If it is false, Replace() will catch any exceptions and continues to write to
            // the remaining files. It will only allow an ArgumentException to bubble up to the caller if SourceFiles is empty.
            // This returns false if something went wrong, and true if all files wrote successfully.

            // If the Aho-Corasick matcher was pre-generated, use that.
            if (_matcher != null)
            {
                return OutputHelper.Replace(
                    ReplacePhrases, SourceFiles, _matcher, wholeWord, preserveCase, throwExceptions, styling);
            }
            return OutputHelper.Replace(
                ReplacePhrases, SourceFiles, wholeWord, caseSensitive, preserveCase, throwExceptions, styling);
        }

        /// <summary>
        /// Asynchronously searches through a list of source files, looking for instances of keys from 
        /// the ReplacePhrases dict, replacing them with the associated value, and then
        /// saving the resulting text off to a list of destination files.
        /// 
        /// If the "throwExceptions" flag is set to false and writing to one of the files failed,
        /// the method will prevent the exception from bubbling up to the caller and continue
        /// to write to the other files in the list.
        /// </summary>
        /// <param name="replacements"></param>
        /// <param name="sourceFiles"></param>
        /// <param name="wholeWord"></param>
        /// <param name="caseSensitive"></param>
        /// <param name="preserveCase"></param>
        /// <param name="styling"></param>
        /// <param name="throwExceptions"></param>
        /// <returns>True is all replacements were made successfully, false if something went wrong.</returns>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="PathTooLongException"></exception>
        /// <exception cref="DirectoryNotFoundException"></exception>
        /// <exception cref="FileNotFoundException"></exception>
        /// <exception cref="IOException"></exception>
        /// <exception cref="UnauthorizedAccessException"></exception>
        /// <exception cref="System.Security.SecurityException"></exception>
        /// <exception cref="InvalidXmlStructureException">
        /// The XML data within a .docx or .xlsx file has an incorrect structure and could not be parsed.
        /// </exception>
        public static async Task<bool> ReplaceAsync(
            Dictionary<string, string> replacements,
            IEnumerable<SourceFile> sourceFiles,
            bool wholeWord = false,
            bool caseSensitive = false,
            bool preserveCase = false,
            Styling? styling = null,
            bool throwExceptions = true)
        {
            // If throwExceptions is true, Replace() will allow exceptions to bubble up to the caller.
            // If it is false, Replace() will catch any exceptions and continues to write to
            // the remaining files. It will only allow an ArgumentException to bubble up to the caller if SourceFiles is empty.
            // This returns false if something went wrong, and true if all files wrote successfully.
            var result = false;
            await Task.Run(() =>
            {
                result = OutputHelper.Replace(
                    replacements, sourceFiles, wholeWord, caseSensitive, preserveCase, throwExceptions, styling);
            });
            return result;
        }

        /// <summary>
        /// Asynchronously earches through a list of source files, looking for instances of keys from 
        /// the ReplacePhrases dict, replacing them with the associated value, and then
        /// saving the resulting text off to a list of destination files.
        /// 
        /// If the "throwExceptions" flag is set to false and writing to one of the files failed,
        /// the method will continue to write to the other files in the list.
        /// </summary>
        /// <param name="wholeWord"></param>
        /// <param name="caseSensitive"></param>
        /// <param name="preserveCase"></param>
        /// <param name="styling"></param>
        /// <param name="throwExceptions"></param>
        /// <returns>True is all replacements were made successfully, false if something went wrong.</returns>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="PathTooLongException"></exception>
        /// <exception cref="DirectoryNotFoundException"></exception>
        /// <exception cref="FileNotFoundException"></exception>
        /// <exception cref="IOException"></exception>
        /// <exception cref="UnauthorizedAccessException"></exception>
        /// <exception cref="System.Security.SecurityException"></exception>
        /// <exception cref="InvalidXmlStructureException">
        /// The XML data within a .docx or .xlsx file has an incorrect structure and could not be parsed.
        /// </exception>
        public async Task<bool> ReplaceAsync(
            bool wholeWord = false,
            bool caseSensitive = false,
            bool preserveCase = false,
            Styling? styling = null,
            bool throwExceptions = true)
        {
            // If throwExceptions is true, Replace() will allow exceptions to bubble up to the caller.
            // If it is false, Replace() will catch any exceptions and continues to write to
            // the remaining files. It will only allow an ArgumentException to bubble up to the caller if SourceFiles is empty.
            // This returns false if something went wrong, and true if all files wrote successfully.
            var result = false;
            await Task.Run(() =>
            {
                // If the Aho-Corasick matcher was pre-generated, use that.
                if (_matcher != null)
                {
                    result = OutputHelper.Replace(
                        ReplacePhrases, SourceFiles, _matcher, wholeWord, preserveCase, throwExceptions, styling);
                }
                result = OutputHelper.Replace(
                    ReplacePhrases, SourceFiles, wholeWord, caseSensitive, preserveCase, throwExceptions, styling);
            });
            return result;
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
        /// <exception cref="InvalidFileTypeException">
        /// A file type is not supported. See documentation for a list of supported file types.
        /// </exception>
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="PathTooLongException"></exception>
        /// <exception cref="DirectoryNotFoundException"></exception>
        /// <exception cref="FileNotFoundException"></exception>
        /// <exception cref="IOException"></exception>
        /// <exception cref="UnauthorizedAccessException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public static Dictionary<string, string> ParseReplacements(string fileName)
        {
            if (File.Exists(fileName) == false)
            {
                throw new FileNotFoundException("Replacements file does not exist");
            }

            if (FileValidation.IsReplaceFileTypeValid(fileName) == false)
            {
                throw new InvalidFileTypeException("Invalid replacements file type: the only supported file types are .xlsx, and UTF-8 text files");
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
        /// <exception cref="InvalidFileTypeException">At least one of the files either does not exist or is not a valid file type.</exception>
        public static List<SourceFile> ZipSourceFiles(
            IEnumerable<string> sourceFileNames,
            IEnumerable<string> outputFileNames,
            bool validateSrcFiles = false)
        {
            if (sourceFileNames.Count() != outputFileNames.Count())
            {
                throw new ArgumentException("The number of source files does not equal the number of output files.");
            }

            if (validateSrcFiles)
            {
                if (AreFilesValid(sourceFileNames, outputFileNames) == false)
                {
                    throw new InvalidFileTypeException("Invalid source file: the only supported file types are .csv, .tsv, .xlsx., .txt, and .text");
                }
            }

            // zip the source file names and the output file names together and then combine the names into SourceFile objects
            return sourceFileNames
                .Zip(outputFileNames, (s, o) => new { SourceFileName = s, OutputFIleName = o })
                .Select(x => new SourceFile(x.SourceFileName, x.OutputFIleName, -1))
                .ToList();
        }

        /// <summary>
        /// Determines whether a replacement file name ends with a valid file type.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>True if the file ends in a valid file type.</returns>
        public static bool IsReplacementFileValid(string fileName)
        {
            return FileValidation.IsReplaceFileTypeValid(fileName);
        }

        /// <summary>
        /// Determines whether an IEnumerable of source files are valid and that the
        /// type conversion to their corresponding outputs file are valid.
        /// </summary>
        /// <param name="sourceFiles"></param>
        /// <param name="outputFiles"></param>
        /// <returns>True if all files end in valid file types.</returns>
        public static bool AreFilesValid(IEnumerable<string> sourceFiles, IEnumerable<string> outputFiles)
        {
            var sourceFilesCount = sourceFiles.Count();
            if (sourceFilesCount != outputFiles.Count())
            {
                return false;
            }

            foreach (var files in sourceFiles.Zip(outputFiles, (source, output) => (source, output)))
            {
                if (AreFilesValid(files.source, files.output) == false)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Determines whether an IEnumerable of source files are valid and that the
        /// type conversion to their corresponding outputs file are valid.
        /// </summary>
        /// <param name="files"></param>
        /// <returns>True if all files end in valid file types.</returns>
        public static bool AreFilesValid(IEnumerable<SourceFile> files)
        {
            foreach (var file in files)
            {
                if (AreFilesValid(file.SourceFileName, file.OutputFileName) == false)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Determines whether a source file is valid and that the
        /// type conversion to its corresponding output file is valid.
        /// </summary>
        /// <param name="sourceFile"></param>
        /// <param name="outputFile"></param>
        /// <returns>True if the file ends in a valid file type.</returns>
        public static bool AreFilesValid(string sourceFile, string outputFile)
        {
            if (FileValidation.IsSourceFileTypeValid(sourceFile) &&
                FileValidation.IsFileConversionValid(sourceFile, outputFile))
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// Determines whether a source file is valid and that the
        /// type conversion to its corresponding output file is valid.
        /// </summary>
        /// <param name="file"></param>
        /// <returns>True if the file ends in a valid file type.</returns>
        public static bool AreFilesValid(SourceFile file)
        {
            if (FileValidation.IsSourceFileTypeValid(file.SourceFileName) &&
                FileValidation.IsFileConversionValid(file.SourceFileName, file.OutputFileName))
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// Generates the Aho-Corasick matcher, which front-loads much of the processing time
        /// rather than performing this operation when the Replace method is called.
        /// </summary>
        /// <param name="replacements"></param>
        /// <param name="caseSensitive"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public void GenerateMatcher(bool caseSensitive)
        {
            if (ReplacePhrases.Count == 0)
            {
                throw new InvalidOperationException("The replacements are empty. Matcher not generated.");
            }

            _matcher = AhoCorasickMatcher.CreateMatcher(ReplacePhrases, caseSensitive);
        }

        /// <summary>
        /// Clears the Aho-Corasick matcher, resulting in it being
        /// generated within the Replace method.
        /// </summary>
        public void ClearMatcher()
        {
            _matcher = null;
        }

        /// <summary>
        /// Checks to see whether a matcher is pre-generated.
        /// </summary>
        /// <returns>Returns true if there is a matcher, false otherwise.</returns>
        public bool IsMatcherCreated()
        {
            return _matcher != null;
        }
    }
}

﻿using TextReplaceAPI.Core.AhoCorasick;
using TextReplaceAPI.Core.Helpers;
using TextReplaceAPI.Core.Validation;
using TextReplaceAPI.Data;
using TextReplaceAPI.Exceptions;

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

        private AhoCorasickStringSearcher? _matcher = null;

        /// <summary>
        /// Initializes a Dictionary<string, string> of replace phrases and an
        /// IEnumberable<SourceFile> containing the source and output file names.
        /// </summary>
        /// <param name="replacementsFileName"></param>
        /// <param name="sourceFileNames"></param>
        /// <param name="outputFileNames"></param>
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
        public Replacify(
            string replacementsFileName,
            IEnumerable<string> sourceFileNames,
            IEnumerable<string> outputFileNames)
		{
            _replacePhrases = ParseReplacements(replacementsFileName);

            if (AreSourceFileTypesValid(sourceFileNames) == false)
            {
                throw new InvalidFileTypeException("Invalid source file: the only supported file types are .csv, .tsv, .xlsx., .txt, and .text");
            }

            if (AreOutputFileTypesValid(sourceFileNames) == false)
            {
                throw new InvalidFileTypeException("Invalid output file: the only supported file types are .csv, .tsv, .xlsx., .txt, and .text");
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
        /// rather than when the PerformReplacements method is called.
        /// </summary>
        /// <param name="replacementsFileName"></param>
        /// <param name="sourceFileNames"></param>
        /// <param name="outputFileNames"></param>
        /// <param name="preGenerateMatcher"></param>
        /// <param name="caseSensitive"></param>
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
        public Replacify(
            string replacementsFileName,
            IEnumerable<string> sourceFileNames,
            IEnumerable<string> outputFileNames,
            bool preGenerateMatcher,
            bool caseSensitive)
        {
            _replacePhrases = ParseReplacements(replacementsFileName);

            if (AreSourceFileTypesValid(sourceFileNames) == false)
            {
                throw new InvalidFileTypeException("Invalid source file: the only supported file types are .csv, .tsv, .xlsx., .txt, and .text");
            }

            if (AreOutputFileTypesValid(sourceFileNames) == false)
            {
                throw new InvalidFileTypeException("Invalid output file: the only supported file types are .csv, .tsv, .xlsx., .txt, and .text");
            }

            // zip the source file names and the output file names together and then combine the names into SourceFile objects
            _sourceFiles = ZipSourceFiles(sourceFileNames, outputFileNames);

            // generate the matcher
            if (preGenerateMatcher)
            {
                GenerateAhoCorasickMatcher(ReplacePhrases, caseSensitive);
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
        public static bool PerformReplacements(
            Dictionary<string, string> replacePhrases,
            IEnumerable<SourceFile> sourceFiles,
            bool wholeWord,
            bool caseSensitive,
            bool preserveCase,
            OutputFileStyling? styling = null,
            bool throwExceptions = true)
        {
            // If throwExceptions is true, PerformReplacements() will allow exceptions to bubble up to the caller.
            // If it is false, PerformReplacements() will catch any exceptions and continues to write to
            // the remaining files. It will only allow an ArgumentException to bubble up to the caller if SourceFiles is empty.
            // This returns false if something went wrong, and true if all files wrote successfully.
            return OutputHelper.PerformReplacements(
                replacePhrases, sourceFiles, wholeWord, caseSensitive, preserveCase, throwExceptions, styling);
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
        public bool PerformReplacements(
            bool wholeWord,
            bool caseSensitive,
            bool preserveCase,
            OutputFileStyling? styling = null,
            bool throwExceptions = true)
        {
            // If throwExceptions is true, PerformReplacements() will allow exceptions to bubble up to the caller.
            // If it is false, PerformReplacements() will catch any exceptions and continues to write to
            // the remaining files. It will only allow an ArgumentException to bubble up to the caller if SourceFiles is empty.
            // This returns false if something went wrong, and true if all files wrote successfully.

            // If the Aho-Corasick matcher was pre-generated, use that.
            if (_matcher != null)
            {
                return OutputHelper.PerformReplacements(
                    ReplacePhrases, SourceFiles, _matcher, wholeWord, preserveCase, throwExceptions, styling);
            }
            return OutputHelper.PerformReplacements(
                ReplacePhrases, SourceFiles, wholeWord, caseSensitive, preserveCase, throwExceptions, styling);
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
            if (FileValidation.IsReplaceFileTypeValid(fileName) == false)
            {
                throw new InvalidFileTypeException($"File type {Path.GetExtension(fileName).ToLower()} is not supported as a replacements file.");
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

        /// <summary>
        /// Generates the Aho-Corasick matcher, which front-loads much of the processing time
        /// rather than performing this operation when the PerformReplacements method is called.
        /// </summary>
        /// <param name="replacePhrases"></param>
        /// <param name="caseSensitive"></param>
        public void GenerateAhoCorasickMatcher(Dictionary<string, string> replacePhrases, bool caseSensitive)
        {
            _matcher = AhoCorasickMatcher.CreateMatcher(ReplacePhrases, caseSensitive);
        }

        /// <summary>
        /// Clears the Aho-Corasick matcher, resulting in it being
        /// generated within the PerformReplacements method.
        /// </summary>
        public void ClearAhoCorasickMatcher()
        {
            _matcher = null;
        }
    }
}

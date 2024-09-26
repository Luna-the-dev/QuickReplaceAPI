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

using TextReplaceAPI.Core.Validation;
using CsvHelper.Configuration;
using CsvHelper;
using System.Globalization;
using System.Text;
using ExcelDataReader;
using System.Data;
using ClosedXML.Excel;

namespace TextReplaceAPI.Core.Helpers
{
    internal static class ReplacementsHelper
    {
        /// <summary>
        /// Parses the replacements from a file. Supports excel files as well as value-seperated files
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>
        /// A dictionary of pairs of the values from the file. If one of the lines in the file has an
        /// incorrect number of values or if the operation fails for another reason, return an empty list.
        /// </returns>
        public static Dictionary<string, string> ParseReplacements(string fileName)
        {
            var phrases = new Dictionary<string, string>();

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using var stream = File.Open(fileName, FileMode.Open, FileAccess.Read);

            using var reader = FileValidation.IsExcelFile(fileName) ?
                ExcelReaderFactory.CreateReader(stream) :
                ExcelReaderFactory.CreateCsvReader(stream);

            while (reader.Read())
            {
                if (reader.GetString(0) == string.Empty)
                {
                    throw new InvalidOperationException("A field within the first column of the replace file is empty.");
                }

                phrases[reader.GetString(0)] = reader.GetString(1);
            }

            if (phrases.Count == 0)
            {
                throw new InvalidOperationException("The dictionary returned by ParseReplacements() is empty.");
            }

            return phrases;
        }

        /// <summary>
        /// Saves the replace phrases list to the file system, performing a sort if requested.
        /// Supported file types are: .csv, .tsv, .xlsx. and .txt
        /// </summary>
        /// <param name="replacePhrases"></param>
        /// <param name="fileName"></param>
        /// <param name="shouldSort"></param>
        /// <param name="delimiter"></param>
        public static void SavePhrasesToFile(Dictionary<string, string> replacePhrases, string fileName, bool shouldSort = false, string delimiter = "")
        {
            if (FileValidation.IsReplaceFileTypeValid(fileName) == false)
            {
                throw new NotSupportedException($"File type {Path.GetExtension(fileName).ToLower()} is not supported.");
            }

            var directory = Path.GetDirectoryName(fileName)?.Replace("\\", "/");
            if (directory == null)
            {
                throw new DirectoryNotFoundException($"Directory was not found: {fileName}");
            }

            if (FileValidation.IsTextFile(Path.GetExtension(fileName)) && DataValidation.IsDelimiterValid(delimiter) == false)
            {
                throw new InvalidOperationException($"Invalid delimiter: {delimiter}");
            }

            // prevents exception if directory doesnt exist
            if (Directory.Exists(directory) == false)
            {
                Directory.CreateDirectory(directory);
            }

            if (FileValidation.IsExcelFile(fileName))
            {
                SavePhrasesToExcel(replacePhrases, fileName, shouldSort);
            }
            else
            {
                SavePhrasesToCsv(replacePhrases, fileName, shouldSort, delimiter);
            }
        }

        /// <summary>
        /// Saves the replace phrases list to the file system as an excel spreadsheet, performing a sort if requested.
        /// </summary>
        /// <param name="replacePhrases"></param>
        /// <param name="fileName"></param>
        /// <param name="shouldSort"></param>
        private static void SavePhrasesToExcel(Dictionary<string, string> replacePhrases, string fileName, bool shouldSort)
        {
            var phrases = shouldSort ? replacePhrases.OrderBy(x => x.Key).ToList() : replacePhrases.ToList();

            using var workbook = new XLWorkbook();
            // limit the length of the worksheet name because excel doesnt allow anything over 31 chars
            string name = Path.GetFileName(fileName);
            if (name.Length > 31)
            {
                name = name.Substring(0, 31);
            }
            var worksheet = workbook.Worksheets.Add(name);

            for (int i = 0; i < phrases.Count; i++)
            {
                worksheet.Cell(i + 1, 1).Value = phrases[i].Key;
                worksheet.Cell(i + 1, 2).Value = phrases[i].Value;
            }

            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(fileName);
        }

        /// <summary>
        /// Saves the replace phrases list to the file system as a delimiter-seperated file, performing a sort if requested.
        /// </summary>
        /// <param name="replacePhrases"></param>
        /// <param name="fileName"></param>
        /// <param name="shouldSort"></param>
        /// <param name="delimiter"></param>
        private static void SavePhrasesToCsv(Dictionary<string, string> replacePhrases, string fileName, bool shouldSort, string delimiter)
        {
            using var writer = new StreamWriter(fileName);
            var csvConfig = new CsvConfiguration(CultureInfo.CurrentCulture)
            {
                Delimiter = delimiter,
                HasHeaderRecord = false,
                Encoding = Encoding.UTF8
            };

            using var csvWriter = new CsvWriter(writer, csvConfig);
            if (shouldSort)
            {
                csvWriter.WriteRecords(replacePhrases.OrderBy(x => x.Key));
            }
            else
            {
                csvWriter.WriteRecords(replacePhrases);
            }
        }
    }
}

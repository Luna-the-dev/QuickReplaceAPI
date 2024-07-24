using System.Diagnostics;
using TextReplaceAPI.Core.Validation;
using CsvHelper.Configuration;
using CsvHelper;
using System.Globalization;
using System.Text;
using ExcelDataReader;
using System.Data;
using ClosedXML.Excel;

namespace TextReplaceAPI.MVVM.Model
{
    internal class ReplaceData
    {
        // key is the phrase to replace, value is what it is being replaced with
        private Dictionary<string, string> _replacePhrases = [];
        public Dictionary<string, string> ReplacePhrases
        {
            get { return _replacePhrases; }
            set { _replacePhrases = value; }
        }

        /// <summary>
        /// Parses through a file for replace phrases and sets the ReplacePhrases List/Dict.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="dryRun"></param>
        /// <returns>True if new replace phrases were successfully set from the file.</returns>
        public bool SetNewReplacePhrasesFromFile(string fileName, bool dryRun = false)
        {
            try
            {
                // check to see if file name is valid so that the phrases can be parsed
                // before setting the file name
                if (FileValidation.IsInputFileReadable(fileName) == false)
                {
                    throw new IOException("Input file is not readable in SetNewReplaceFileFromUser().");
                }

                // if caller specified that this should be a dry run,
                // then dont actually assign the parsed data to the dict
                if (dryRun)
                {
                    // will throw InvalidOperationException if it returns a dict of count == 0
                    ParseReplacements(fileName);
                    return true;
                }

                // parse through phrases and attempt to save them
                ReplacePhrases = ParseReplacements(fileName);

                return true;
            }
            catch (IOException e)
            {
                Debug.WriteLine(e);
                return false;
            }
            catch (InvalidOperationException e)
            {
                Debug.WriteLine(e);
                return false;
            }
            catch (NotSupportedException e)
            {
                Debug.WriteLine(e);
                return false;
            }
            catch (CsvHelper.MissingFieldException)
            {
                Debug.WriteLine("CsvHelper could not parse the file with the given delimiter.");
                return false;
            }
            catch
            {
                Debug.WriteLine("Something unexpected happened in SetNewReplaceFileFromUser().");
                return false;
            }
        }

        /// <summary>
        /// Parses the replacements from a file. Supports excel files as well as value-seperated files
        /// Supported file types are: .csv, .tsv, .xlsx. and .txt
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>
        /// A dictionary of pairs of the values from the file. If one of the lines in the file has an
        /// incorrect number of values or if the operation fails for another reason, return an empty list.
        /// </returns>
        public static Dictionary<string, string> ParseReplacements(string fileName)
        {
            if (FileValidation.IsReplaceFileTypeValid(fileName) == false)
            {
                throw new NotSupportedException($"File type {Path.GetExtension(fileName).ToLower()} is not supported.");
            }

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
        /// <param name="fileName"></param>
        /// <param name="shouldSort"></param>
        /// <param name="delimiter"></param>
        public void SavePhrasesToFile(string fileName, bool shouldSort = false, string delimiter = "")
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
                SavePhrasesToExcel(fileName, shouldSort);
            }
            else
            {
                SavePhrasesToCsv(fileName, shouldSort, delimiter);
            }
        }

        private void SavePhrasesToExcel(string fileName, bool shouldSort)
        {
            var phrases = (shouldSort) ? ReplacePhrases.OrderBy(x => x.Key).ToList() : ReplacePhrases.ToList();

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

        private void SavePhrasesToCsv(string fileName, bool shouldSort, string delimiter)
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
                csvWriter.WriteRecords(ReplacePhrases.OrderBy(x => x.Key));
            }
            else
            {
                csvWriter.WriteRecords(ReplacePhrases);
            }
        }
    }
}

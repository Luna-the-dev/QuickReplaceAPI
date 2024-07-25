using CsvHelper.Configuration.Attributes;

namespace TextReplaceAPI.Data
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

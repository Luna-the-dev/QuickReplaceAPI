namespace TextReplaceAPI.DataTypes
{
    public record SourceFile
    {
        public string SourceFileName { get; set; }
        public string OutputFileName { get; set; }
        public int NumOfReplacements { get; set; }

        public SourceFile(
            string sourceFileName,
            string outputFileName,
            int numOfReplacements)
        {
            SourceFileName = sourceFileName;
            OutputFileName = outputFileName;
            NumOfReplacements = numOfReplacements;
        }
    }
}

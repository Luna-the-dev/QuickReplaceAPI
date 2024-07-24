namespace TextReplaceAPI.Core.Enums
{
    internal enum OutputFileTypeEnum
    {
        KeepFileType,
        Document,
        Text
    }

    internal class OutputFileTypeClass
    {
        public static string OutputFileTypeString(OutputFileTypeEnum fileType, string fileName)
        {
            return fileType switch
            {
                OutputFileTypeEnum.KeepFileType => Path.GetExtension(fileName),
                OutputFileTypeEnum.Text => ".txt",
                OutputFileTypeEnum.Document => ".docx",
                _ => throw new NotImplementedException($"{fileType} is not implemented in OutputFileTypeString()")
            };
        }
    }

}

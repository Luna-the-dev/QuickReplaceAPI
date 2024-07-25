namespace TextReplaceAPI.Exceptions
{
    public class InvalidXmlStructureException : Exception
    {
        public InvalidXmlStructureException()
        {
        }

        public InvalidXmlStructureException(string message) : base(message)
        {
        }

        public InvalidXmlStructureException(string message, Exception inner) : base(message, inner)
        {
        }
    }
}

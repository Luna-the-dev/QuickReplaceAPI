namespace TextReplaceAPI.Exceptions
{
    public class InvalidFileTypeException : Exception
    {
        public InvalidFileTypeException()
        {
        }

        public InvalidFileTypeException(string message) : base(message)
        {
        }

        public InvalidFileTypeException(string message, Exception inner) : base(message, inner)
        {
        }
    }
}

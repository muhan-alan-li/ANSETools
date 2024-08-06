namespace ANSExlsx2txt.Exceptions;

public class CellParsingException : Exception
{
    public CellParsingException()
    {
    }
    
    public CellParsingException(string message)
        : base(message)
    {
    }

    public CellParsingException(string message, Exception inner)
        : base(message, inner)
    {
    }
}
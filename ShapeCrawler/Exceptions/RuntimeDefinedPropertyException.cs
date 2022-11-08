namespace ShapeCrawler.Exceptions;

internal class RuntimeDefinedPropertyException : ShapeCrawlerException
{
    public RuntimeDefinedPropertyException(string message)
        : base(message, ExceptionCode.RuntimeDefinedPropertyException)
    {
    }
}
namespace ShapeCrawler.Exceptions;

internal class RuntimeDefinedPropertyException : ShapeCrawlerException
{
    internal RuntimeDefinedPropertyException(string message)
        : base(message, ExceptionCode.RuntimeDefinedPropertyException)
    {
    }
}
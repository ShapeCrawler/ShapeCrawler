namespace ShapeCrawler.Exceptions;

internal sealed class RuntimeDefinedPropertyException : ShapeCrawlerException
{
    internal RuntimeDefinedPropertyException(string message)
        : base(message, ExceptionCode.RuntimeDefinedPropertyException)
    {
    }
}
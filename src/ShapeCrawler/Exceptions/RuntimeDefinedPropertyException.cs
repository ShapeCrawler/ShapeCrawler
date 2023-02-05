namespace ShapeCrawler.Exceptions;

internal sealed class RuntimeDefinedPropertyException : SCException
{
    internal RuntimeDefinedPropertyException(string message)
        : base(message, ExceptionCode.RuntimeDefinedPropertyException)
    {
    }
}
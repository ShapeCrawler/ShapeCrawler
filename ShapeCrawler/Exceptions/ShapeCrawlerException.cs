using System;

namespace ShapeCrawler.Exceptions;

internal class ShapeCrawlerException : Exception
{
    internal ShapeCrawlerException()
    {
    }

    internal ShapeCrawlerException(string message)
        : base(message + "\nIf you have a question, feel free to report an issue https://github.com/ShapeCrawler/ShapeCrawler/issues")
    {
    }

    internal ShapeCrawlerException(string message, int errorCode)
        : base(message)
    {
    }

    internal ShapeCrawlerException(string message, ExceptionCode exceptionCode)
        : base(message)
    {
    }

    internal ShapeCrawlerException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
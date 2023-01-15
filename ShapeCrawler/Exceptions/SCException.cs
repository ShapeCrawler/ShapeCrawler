using System;

namespace ShapeCrawler.Exceptions;

internal class SCException : Exception
{
    internal SCException()
    {
    }

    internal SCException(string message)
        : base(message + "\nIf you have a question, feel free to report an issue https://github.com/ShapeCrawler/ShapeCrawler/issues")
    {
    }

    internal SCException(string message, int errorCode)
        : base(message)
    {
    }

    internal SCException(string message, ExceptionCode exceptionCode)
        : base(message)
    {
    }

    internal SCException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
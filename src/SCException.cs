using System;

namespace ShapeCrawler;

/// <summary>
///     Represents an exception that is thrown when a ShapeCrawler error occurs.
/// </summary>
public sealed class SCException : Exception
{
    internal SCException()
    {
    }

    internal SCException(string message)
        : base($"{message}{Environment.NewLine}{Environment.NewLine}If you have a question, feel free to report an issue https://github.com/ShapeCrawler/ShapeCrawler/issues")
    {
    }

    internal SCException(string message, Exception innerException)
        : base(
            $"{message}{Environment.NewLine}{Environment.NewLine}If you have a question, feel free to report an issue https://github.com/ShapeCrawler/ShapeCrawler/issues",
            innerException)
    {
    }
}
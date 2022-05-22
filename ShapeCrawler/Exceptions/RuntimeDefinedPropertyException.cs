using System;

namespace ShapeCrawler.Exceptions
{
    /// <summary>
    ///     Thrown while attempting to access runtime defined property, but it does not exist for the current object.
    /// </summary>
    internal class RuntimeDefinedPropertyException : ShapeCrawlerException
    {
        public RuntimeDefinedPropertyException(string message)
            : base(message, ExceptionCode.RuntimeDefinedPropertyException)
        {
        }
    }
}
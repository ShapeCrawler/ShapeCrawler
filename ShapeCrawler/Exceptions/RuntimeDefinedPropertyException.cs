using System;

namespace ShapeCrawler.Exceptions
{
    /// <summary>
    ///     Thrown while attempting to access runtime defined property, but it does not exist for the current object.
    /// </summary>
    internal class RuntimeDefinedPropertyException : ShapeCrawlerException
    {
        #region Constructors

        public RuntimeDefinedPropertyException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        public RuntimeDefinedPropertyException(string message)
            : base(message, ExceptionCode.RuntimeDefinedPropertyException)
        {
        }

        public RuntimeDefinedPropertyException()
        {
        }

        #endregion Constructors
    }
}
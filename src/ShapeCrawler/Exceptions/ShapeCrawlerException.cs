using System;
using ShapeCrawler.Enums;

namespace ShapeCrawler.Exceptions
{
    /// <summary>
    /// Represents the library exception. 
    /// </summary>
    public class ShapeCrawlerException : Exception
    {
        #region Properties

        /// <summary>
        /// Returns error code number.
        /// </summary>
        public int ErrorCode { get; } = (int)ExceptionCode.ShapeCrawlerException;

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Defines a parameterless constructor.
        /// </summary>
        public ShapeCrawlerException() { }

        public ShapeCrawlerException(string message) : base(message) { }

        public ShapeCrawlerException(string message, int errorCode) : base(message)
        {
            ErrorCode = errorCode;
        }

        public ShapeCrawlerException(string message, ExceptionCode exceptionCode) : base(message)
        {
            ErrorCode = (int)exceptionCode;
        }

        public ShapeCrawlerException(int errorCode)
        {
            ErrorCode = errorCode;
        }

        public ShapeCrawlerException(string message, Exception innerException) : base(message, innerException)
        {
        }

        #endregion Constructors
    }
}

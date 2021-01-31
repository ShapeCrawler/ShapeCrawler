using System;

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

        public ShapeCrawlerException() { }

        public ShapeCrawlerException(string message)
        {
            message +=
                "\nIf you have a question, feel free to report the issue https://github.com/ShapeCrawler/ShapeCrawler/issues";
            
        }

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

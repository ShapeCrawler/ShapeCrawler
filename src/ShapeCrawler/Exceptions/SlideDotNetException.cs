using System;
using ShapeCrawler.Enums;

namespace ShapeCrawler.Exceptions
{
    /// <summary>
    /// Represents the library exception. 
    /// </summary>
    public class SlideDotNetException : Exception
    {
        #region Properties

        /// <summary>
        /// Returns error code number.
        /// </summary>
        public int ErrorCode { get; } = (int)ExceptionCodes.SlideDotNetException;

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Defines a parameterless constructor.
        /// </summary>
        public SlideDotNetException() { }

        public SlideDotNetException(string message) : base(message) { }

        public SlideDotNetException(string message, int errorCode) : base(message)
        {
            ErrorCode = errorCode;
        }

        public SlideDotNetException(string message, ExceptionCodes exceptionCode) : base(message)
        {
            ErrorCode = (int)exceptionCode;
        }

        public SlideDotNetException(int errorCode)
        {
            ErrorCode = errorCode;
        }

        public SlideDotNetException(string message, Exception innerException) : base(message, innerException)
        {
        }

        #endregion Constructors
    }
}

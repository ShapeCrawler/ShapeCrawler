using System;
using SlideDotNet.Enums;

namespace SlideDotNet.Exceptions
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
        public int ErrorCode { get; } = (int)ExceptionCodes.SlideXmlException;

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

        public SlideDotNetException(int errorCode)
        {
            ErrorCode = errorCode;
        }

        #endregion Constructors
    }
}

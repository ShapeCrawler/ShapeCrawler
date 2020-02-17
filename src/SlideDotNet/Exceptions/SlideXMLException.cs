using System;
using SlideDotNet.Enums;

namespace SlideDotNet.Exceptions
{
    /// <summary>
    /// Represents the library exception. 
    /// </summary>
    public class SlideXmlException : Exception
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
        public SlideXmlException() { }

        public SlideXmlException(string message) : base(message) { }

        public SlideXmlException(string message, int errorCode) : base(message)
        {
            ErrorCode = errorCode;
        }

        public SlideXmlException(int errorCode)
        {
            ErrorCode = errorCode;
        }

        #endregion Constructors
    }
}

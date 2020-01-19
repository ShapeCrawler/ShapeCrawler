using System;

namespace SlideXML.Exceptions
{
    /// <summary>
    /// Represents the library exception. 
    /// </summary>
    public class SlideXMLException : Exception
    {
        #region Properties

        /// <summary>
        /// Returns error code number.
        /// </summary>
        public int ErrorCode { get; } = 100; // 100 is general code

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Defines parametersless constructor.
        /// </summary>
        public SlideXMLException() { }

        public SlideXMLException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="SlideXMLException"/> class with default error message.
        /// </summary>
        public SlideXMLException(int errorCode, string message) : base(message)
        {
            ErrorCode = errorCode;
        }

        #endregion Constructors
    }
}

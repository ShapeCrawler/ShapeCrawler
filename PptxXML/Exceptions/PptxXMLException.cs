using System;

namespace PptxXML.Exceptions
{
    /// <summary>
    /// Represents the library exception. 
    /// </summary>
    public class PptxXMLException : Exception
    {
        #region Properties

        /// <summary>
        /// Returns error code number.
        /// </summary>
        public int ErrorCode { get; } = 100; // 100 is general code

        #endregion Properties

        /// <summary>
        /// Initializes a new instance of the <see cref="PptxXMLException"/> class with default error message.
        /// </summary>
        public PptxXMLException(int errorCode, string message) : base(message)
        {
            ErrorCode = errorCode;
        }

        /// <summary>
        /// Defines parametersless constructor.
        /// </summary>
        public PptxXMLException() { }
    }
}

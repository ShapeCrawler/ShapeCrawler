using System;

namespace PptxXML.Exceptions
{
    /// <summary>
    /// Represent an exception which is throws when a type of element could not be defined.
    /// </summary>
    public class TypeException : Exception
    {
        #region Fields

        private const string DefaultMessage = "Element type was not defined";

        #endregion

        #region Constructors

        /// <summary>
        /// Initialise <see cref="TypeException"/> exception with default error message.
        /// </summary>
        public TypeException():
            base(DefaultMessage)
        {
            
        }

        #endregion
    }
}

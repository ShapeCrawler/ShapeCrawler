using SlideDotNet.Enums;

namespace SlideDotNet.Exceptions
{
    /// <summary>
    /// Thrown while attempting to access runtime defined property, but it does not exist for the current object.
    /// </summary>
    public class RuntimeDefinedPropertyException : SlideDotNetException
    {
        #region Constructors

        public RuntimeDefinedPropertyException(string message) 
            : base(message, ExceptionCodes.RuntimeDefinedPropertyException) { }

        #endregion Constructors
    }
}

using SlideDotNet.Enums;

namespace SlideDotNet.Exceptions
{
    /// <summary>
    /// Thrown when a feature not yet been implemented.
    /// </summary>
    public class NextVersionFeatureException : SlideDotNetException
    {
        public NextVersionFeatureException(string message) 
            : base(message, ExceptionCodes.NextVersionFeatureException) { }
    }
}

using SlideDotNet.Enums;

namespace SlideDotNet.Exceptions
{
    /// <summary>
    /// Thrown when a feature not yet been implemented.
    /// </summary>
    public class FeatureNotYetImplementedException : SlideDotNetException
    {
        public FeatureNotYetImplementedException() : base((int)ExceptionCodes.FeatureNotYetImplementedException)
        {

        }
    }
}

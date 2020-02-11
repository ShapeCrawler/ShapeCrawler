using SlideXML.Enums;

namespace SlideXML.Exceptions
{
    /// <summary>
    /// Thrown when a feature not yet been implemented.
    /// </summary>
    public class FeatureNotYetImplementedException : SlideXmlException
    {
        public FeatureNotYetImplementedException() : base((int)ExceptionCodes.FeatureNotYetImplementedException)
        {

        }
    }
}

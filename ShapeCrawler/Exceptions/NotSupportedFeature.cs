using System;

namespace ShapeCrawler.Exceptions
{
    /// <summary>
    ///     Thrown when the feature has not yet been supported.
    /// </summary>
    internal class NotSupportedFeature : ShapeCrawlerException
    {
        public NotSupportedFeature(string message)
            : base(message, ExceptionCode.NextVersionFeatureException)
        {
        }

        public NotSupportedFeature()
        {
        }

        public NotSupportedFeature(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
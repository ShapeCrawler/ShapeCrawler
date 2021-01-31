namespace ShapeCrawler.Exceptions
{
    /// <summary>
    /// Thrown when a feature not yet been implemented.
    /// </summary>
    public class NextVersionFeatureException : ShapeCrawlerException
    {
        public NextVersionFeatureException(string message) 
            : base(message, ExceptionCode.NextVersionFeatureException) { }

        public NextVersionFeatureException()
        {
        }

        public NextVersionFeatureException(string message, System.Exception innerException) : base(message, innerException)
        {
        }
    }
}

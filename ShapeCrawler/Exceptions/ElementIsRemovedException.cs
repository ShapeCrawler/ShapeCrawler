namespace ShapeCrawler.Exceptions
{
    /// <summary>
    ///     Thrown in access to property of removed presentation element.
    /// </summary>
    public class ElementIsRemovedException : ShapeCrawlerException
    {
        public ElementIsRemovedException(string message, System.Exception innerException) : base(message, innerException)
        {
        }

        public ElementIsRemovedException(string message) : base(message)
        {
        }

        public ElementIsRemovedException()
        {
        }
    }
}
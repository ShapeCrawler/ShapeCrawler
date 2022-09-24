namespace ShapeCrawler.Exceptions
{
    /// <summary>
    ///     Thrown when access to removed presentation element.
    /// </summary>
    public class ElementIsRemovedException : ShapeCrawlerException
    {
        internal ElementIsRemovedException(string message)
            : base(message)
        {
        }
    }
}
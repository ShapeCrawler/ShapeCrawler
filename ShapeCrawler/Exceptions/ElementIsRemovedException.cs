namespace ShapeCrawler.Exceptions
{
    /// <summary>
    ///     Thrown when access to removed presentation element.
    /// </summary>
    internal class ElementIsRemovedException : ShapeCrawlerException
    {
        internal ElementIsRemovedException(string message)
            : base(message)
        {
        }
    }
}
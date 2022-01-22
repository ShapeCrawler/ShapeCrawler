namespace ShapeCrawler.Exceptions
{
    /// <summary>
    ///     Thrown when access to removed presentation element.
    /// </summary>
    public class ElementIsRemovedException : ShapeCrawlerException
    {
        /// <summary>
        ///     Initializes a new instance of the <see cref="ElementIsRemovedException"/> class.
        /// </summary>
        internal ElementIsRemovedException(string message)
            : base(message)
        {
        }
    }
}
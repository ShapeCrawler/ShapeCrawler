namespace ShapeCrawler.Exceptions
{
    internal class PlaceholderCannotBeChangedException : ShapeCrawlerException
    {
        internal PlaceholderCannotBeChangedException()
            : base("The shape is a placeholder and cannot be changed on the Slide level")
        {
        }
    }
}
namespace ShapeCrawler.Exceptions
{
    public class PlaceholderCannotBeChangedException : ShapeCrawlerException
    {
        internal PlaceholderCannotBeChangedException()
            : base("The shape is a placeholder and cannot be changed on the slide level")
        {
        }
    }
}
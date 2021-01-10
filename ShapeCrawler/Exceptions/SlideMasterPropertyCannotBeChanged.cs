namespace ShapeCrawler.Exceptions
{
    internal class SlideMasterPropertyCannotBeChanged : ShapeCrawlerException
    {
        public SlideMasterPropertyCannotBeChanged()
        {
        }

        public SlideMasterPropertyCannotBeChanged(string message) : base(message)
        {
        }

        public SlideMasterPropertyCannotBeChanged(string message, System.Exception innerException) : base(message, innerException)
        {
        }
    }
}
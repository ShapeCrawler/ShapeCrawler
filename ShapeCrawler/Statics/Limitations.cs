namespace ShapeCrawler.Statics
{
    /// <summary>
    ///     Contains library limitation..
    /// </summary>
    public static class Limitations
    {
        /// <summary>
        ///     Gets the maximal allowed size of presentation in bytes.
        /// </summary>
        public static int MaxPresentationSize => 250 * 1024 * 1024;

        /// <summary>
        ///     Gets the maximal allowed number of slides in a presentation.
        /// </summary>
        public static int MaxSlidesNumber => 300;
    }
}
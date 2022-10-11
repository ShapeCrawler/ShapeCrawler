namespace ShapeCrawler.Constants
{
    /// <summary>
    ///     Contains library limitation.
    /// </summary>
    internal static class Limitations
    {
        /// <summary>
        ///     Gets the maximal allowed size of presentation in bytes.
        /// </summary>
        public static int MaxPresentationSize => 250 * 1024 * 1024;

        /// <summary>
        ///     Gets the maximal allowed number of slides in a presentation.
        /// </summary>
        internal static int MaxSlidesNumber => 300;
    }
}
namespace ShapeCrawler.Statics
{
    /// <summary>
    ///     Contains library limitation constant values.
    /// </summary>
    public static class Limitations
    {
        /// <summary>
        ///     Gets the maximal allowed size of presentation in bytes.
        /// </summary>
        public static int MaxPresentationSize => 157286400; // 150 MB

        /// <summary>
        ///     Gets the maximal allowed number of slides in a presentation.
        /// </summary>
        public static int MaxSlidesNumber => 250;
    }
}
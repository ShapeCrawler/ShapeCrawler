namespace SlideDotNet
{
    /// <summary>
    /// Contains library limitation constant values.
    /// </summary>
    public static class Limitations
    {
        /// <summary>
        /// Returns the maximal allowed size of presentation in bytes.
        /// </summary>
        public static int MaxPresentationSize => 157286400; // 150 MB

        /// <summary>
        /// Returns the maximal allowed number of slides in a presentation.
        /// </summary>
        public static int MaxSlidesNumber => 250;

        /// <summary>
        /// Returns the maximal allowed number of shapes on one slide.
        /// </summary>
        public static int MaxShapesNumber => 100;
    }
}

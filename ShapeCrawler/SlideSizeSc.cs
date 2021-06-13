namespace ShapeCrawler.Models
{
    /// <summary>
    ///     Represent presentation slides size.
    /// </summary>
    internal class SlideSizeSc
    {
        #region Constructors

        public SlideSizeSc(int slideWidth, int slideHeight)
        {
            Width = slideWidth;
            Height = slideHeight;
        }

        #endregion Constructors

        #region Properties

        public int Width { get; }

        public int Height { get; }

        #endregion Properties
    }
}
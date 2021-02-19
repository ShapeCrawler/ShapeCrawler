namespace ShapeCrawler.Models.SlideComponents
{
    /// <summary>
    ///     Represent presentation slides size.
    /// </summary>
    public class SlideSizeSc
    {
        #region Constructors

        public SlideSizeSc(int sdkW, int sdkH)
        {
            Width = sdkW;
            Height = sdkH;
        }

        #endregion Constructors

        #region Properties

        public int Width { get; }

        public int Height { get; }

        #endregion Properties
    }
}
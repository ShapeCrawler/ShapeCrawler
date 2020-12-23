namespace ShapeCrawler.Models
{
    /// <summary>
    /// Represent presentation slides size.
    /// </summary>
    public class SlideSize
    {
        #region Properties

        public int Width { get; }

        public int Height { get; }

        #endregion Properties

        #region Constructors

        public SlideSize(int sdkW, int sdkH)
        {
            Width = sdkW;
            Height = sdkH;
        }

        #endregion Constructors
    }
}
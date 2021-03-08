using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing
{
    /// <summary>
    ///     Represents a picture on a Slide Layout.
    /// </summary>
    internal class LayoutPicture : LayoutShape, IShape
    {
        internal LayoutPicture(SlideLayoutSc slideLayout, P.Picture pPicture)
            : base(slideLayout, pPicture)
        {
        }
    }
}
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing
{
    /// <summary>
    ///     Represents a picture on a Slide Layout.
    /// </summary>
    internal class LayoutPicture : LayoutShape, IShape
    {
        internal LayoutPicture(SCSlideLayout slideInternalLayout, P.Picture pPicture)
            : base(slideInternalLayout, pPicture)
        {
        }

        public override SCSlideMaster ParentSlideMaster { get; set; }
    }
}
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a Slide Layout.
    /// </summary>
    public abstract class LayoutShape : Shape
    {
        protected LayoutShape(SlideLayoutSc slideLayout, OpenXmlCompositeElement pShapeTreeChild) : base(
            pShapeTreeChild)
        {
            SlideLayout = slideLayout;
        }

        internal SlideLayoutSc SlideLayout { get; }

        public override IPlaceholder Placeholder => LayoutPlaceholder.Create(PShapeTreeChild, this);
        public override ThemePart ThemePart => SlideLayout.SlideLayoutPart.SlideMasterPart.ThemePart;
        public override PresentationSc Presentation => SlideLayout.SlideMaster.Presentation;
    }
}
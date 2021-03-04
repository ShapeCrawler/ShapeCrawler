using DocumentFormat.OpenXml;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a Slide Layout.
    /// </summary>
    public abstract class LayoutShape : Shape
    {
        protected LayoutShape(OpenXmlCompositeElement pShapeTreeChild, SlideSc slide) : base(pShapeTreeChild, slide)
        {

        }

        public override IPlaceholder Placeholder => LayoutPlaceholder.Create(PShapeTreeChild, this);

        public SlideLayoutSc SlideLayout { get; internal set; }
    }
}
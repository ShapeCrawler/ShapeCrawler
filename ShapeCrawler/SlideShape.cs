using DocumentFormat.OpenXml;
using ShapeCrawler.Placeholders;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a Slide.
    /// </summary>
    public abstract class SlideShape : Shape
    {
        protected SlideShape(OpenXmlCompositeElement pShapeTreeChild, SlideSc slide) : base(pShapeTreeChild, slide)
        {
        }

        public override IPlaceholder Placeholder => SlidePlaceholder.Create(PShapeTreeChild, this);
    }
}
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Placeholders;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a Slide.
    /// </summary>
    public abstract class SlideShape : Shape
    {
        internal SlideSc Slide { get; }

        protected SlideShape(SlideSc slide, OpenXmlCompositeElement pShapeTreeChild) : base(pShapeTreeChild)
        {
            Slide = slide;
        }

        public override IPlaceholder Placeholder => SlidePlaceholder.Create(PShapeTreeChild, this);

        public override ThemePart ThemePart => Slide.SlidePart.SlideLayoutPart.SlideMasterPart.ThemePart;
        public override PresentationSc Presentation => Slide.Presentation;
    }
}
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a Slide.
    /// </summary>
    public abstract class SlideShape : Shape
    {
        protected SlideShape(SlideSc slide, OpenXmlCompositeElement pShapeTreeChild) : base(pShapeTreeChild)
        {
            Slide = slide;
        }

        internal SlideSc Slide { get; }

        public override IPlaceholder Placeholder => SlidePlaceholder.Create(PShapeTreeChild, this);

        internal override ThemePart ThemePart => Slide.SlidePart.SlideLayoutPart.SlideMasterPart.ThemePart;
        public override PresentationSc Presentation => Slide.Presentation;
        public override SlideMasterSc SlideMaster => Slide.SlideLayout.SlideMaster;
    }
}
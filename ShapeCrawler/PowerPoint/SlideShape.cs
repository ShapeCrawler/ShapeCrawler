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
        protected SlideShape(SCSlide slide, OpenXmlCompositeElement pShapeTreeChild) : base(pShapeTreeChild)
        {
            Slide = slide;
        }

        internal SCSlide Slide { get; }

        public override IPlaceholder Placeholder => SlidePlaceholder.Create(PShapeTreeChild, this);

        internal override ThemePart ThemePart => Slide.SlidePart.SlideLayoutPart.SlideMasterPart.ThemePart;
        public override SCPresentation Presentation => Slide.Presentation;
        public override SCSlideMaster SlideMaster => Slide.SlideLayout.SlideMaster;
    }
}
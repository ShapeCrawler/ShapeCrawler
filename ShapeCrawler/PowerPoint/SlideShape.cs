using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Placeholders;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a Slide.
    /// </summary>
    internal abstract class SlideShape : Shape
    {
        internal override ThemePart ThemePart => Slide.SlidePart.SlideLayoutPart.SlideMasterPart.ThemePart;
        internal SCSlide Slide { get; }

        protected SlideShape(SCSlide slide, OpenXmlCompositeElement pShapeTreeChild) : base(pShapeTreeChild)
        {
            Slide = slide;
        }

        public override IPlaceholder Placeholder => SlidePlaceholder.Create(PShapeTreeChild, this);
        public override SCPresentation Presentation => Slide.Presentation;
        public override SCSlideMaster SlideMaster => Slide.SlideLayout.SlideMaster;
    }
}
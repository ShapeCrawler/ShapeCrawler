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
        protected SlideShape(SCSlide slide, OpenXmlCompositeElement pShapeTreeChild)
            : base(pShapeTreeChild)
        {
            this.Slide = slide;
        }

        #region Public Properties

        public override IPlaceholder Placeholder => SlidePlaceholder.Create(this.PShapeTreeChild, this);

        public override SCPresentation ParentPresentation => Slide.ParentPresentation;

        public override SCSlideMaster SlideMaster => Slide.Layout.SlideMaster;

        #endregion Public Properties

        internal override ThemePart ThemePart => this.Slide.SlidePart.SlideLayoutPart.SlideMasterPart.ThemePart;

        internal SCSlide Slide { get; }

        
    }
}
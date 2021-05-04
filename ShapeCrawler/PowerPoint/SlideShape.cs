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

        public override SCPresentation ParentPresentation => this.Slide.ParentPresentation;

        public override SCSlideMaster ParentSlideMaster => (SCSlideMaster)this.Slide.ParentSlideLayout.ParentSlideMaster;

        #endregion Public Properties

        internal SCSlide Slide { get; }
    }
}
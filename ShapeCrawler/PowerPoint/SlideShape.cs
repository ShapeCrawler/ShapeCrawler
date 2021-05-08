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
        protected SlideShape(SCSlide parentSlide, OpenXmlCompositeElement sdkPShapeTreeChild)
            : base(sdkPShapeTreeChild, parentSlide)
        {
            this.ParentSlide = parentSlide;
        }

        #region Public Properties

        public override IPlaceholder Placeholder => SlidePlaceholder.Create(this.SdkPShapeTreeChild, this);

        public override SCPresentation ParentPresentation => this.ParentSlide.ParentPresentation;

        public override SCSlideMaster ParentSlideMaster => (SCSlideMaster)this.ParentSlide.ParentSlideLayout.ParentSlideMaster;

        #endregion Public Properties

        internal SCSlide ParentSlide { get; }
    }
}
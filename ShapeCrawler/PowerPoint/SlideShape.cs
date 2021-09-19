using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a Slide.
    /// </summary>
    internal abstract class SlideShape : Shape, IPresentationComponent
    {
        protected SlideShape(SCSlide parentSlide, OpenXmlCompositeElement sdkPShapeTreeChild)
            : base(sdkPShapeTreeChild, parentSlide)
        {
            this.ParentSlide = parentSlide;
        }

        public override IPlaceholder Placeholder => SlidePlaceholder.Create(this.SdkPShapeTreeChild, this);

        public override SCSlideMaster ParentSlideMaster => (SCSlideMaster)this.ParentSlide.ParentSlideLayout.ParentSlideMaster;

        public SCPresentation ParentPresentation => this.ParentSlide.ParentPresentation;

        internal SCSlide ParentSlide { get; }
    }
}
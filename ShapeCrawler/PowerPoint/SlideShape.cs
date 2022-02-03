using DocumentFormat.OpenXml;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape located on Slide.
    /// </summary>
    internal abstract class SlideShape : Shape, IPresentationComponent
    {
        protected SlideShape(OpenXmlCompositeElement childOfpShapeTree, SCSlide slide, Shape groupShape)
            : base(childOfpShapeTree, slide, groupShape)
        {
            this.Slide = slide;
        }

        protected SlideShape(OpenXmlCompositeElement childOfpShapeTree, SCSlide slide)
            : base(childOfpShapeTree, slide)
        {
            this.Slide = slide;
        }

        public override IPlaceholder Placeholder => SlidePlaceholder.Create(this.PShapeTreesChild, this);

        internal override SCSlideMaster SlideMasterInternal
        {
            get => (SCSlideMaster)this.Slide.ParentSlideLayout.ParentSlideMaster;

            set
            {
            }
        }

        public SCPresentation PresentationInternal => this.Slide.parentPresentationInternal;

        public ISlide ParentSlide => this.Slide;

        internal SCSlide Slide { get; }
    }
}
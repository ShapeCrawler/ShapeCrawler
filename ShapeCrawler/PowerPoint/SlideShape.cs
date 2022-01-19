using DocumentFormat.OpenXml;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents shape located on slide.
    /// </summary>
    internal abstract class SlideShape : Shape, IPresentationComponent
    {
        protected SlideShape(OpenXmlCompositeElement pShapeTreesChild, SCSlide parentSlideLayoutInternal, Shape parentGroupShape)
            : base(pShapeTreesChild, parentSlideLayoutInternal, parentGroupShape)
        {
            this.ParentSlideLayoutInternal = parentSlideLayoutInternal;
        }

        public override IPlaceholder Placeholder => SlidePlaceholder.Create(this.PShapeTreesChild, this);

        public override SCSlideMaster ParentSlideMaster
        {
            get => (SCSlideMaster)this.ParentSlideLayoutInternal.ParentSlideLayout.ParentSlideMaster;

            set
            {
            }
        }

        public SCPresentation ParentPresentationInternal => this.ParentSlideLayoutInternal.parentPresentationInternal;

        public ISlide ParentSlide => this.ParentSlideLayoutInternal;

        internal SCSlide ParentSlideLayoutInternal { get; }
    }
}
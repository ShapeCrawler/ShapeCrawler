using DocumentFormat.OpenXml;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a slide.
    /// </summary>
    internal abstract class SlideShape : Shape, IPresentationComponent
    {
        protected SlideShape(OpenXmlCompositeElement pShapeTreesChild, SCSlide parentSlideInternal, Shape parentGroupShape)
            : base(pShapeTreesChild, parentSlideInternal, parentGroupShape)
        {
            this.ParentSlideInternal = parentSlideInternal;
        }

        public override IPlaceholder Placeholder => SlidePlaceholder.Create(this.PShapeTreesChild, this);

        public override SCSlideMaster ParentSlideMaster
        {
            get => (SCSlideMaster)this.ParentSlideInternal.ParentSlideLayout.ParentSlideMaster;

            set
            {
            }
        }

        public SCPresentation ParentPresentationInternal => this.ParentSlideInternal.parentPresentationInternal;

        public ISlide ParentSlide => this.ParentSlideInternal;

        internal SCSlide ParentSlideInternal { get; }
    }
}
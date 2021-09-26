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
        protected SlideShape(OpenXmlCompositeElement pShapeTreesChild, SCSlide parentSlide, Shape parentGroupShape)
            : base(pShapeTreesChild, parentSlide, parentGroupShape)
        {
            this.ParentSlide = parentSlide;
        }

        public override IPlaceholder Placeholder => SlidePlaceholder.Create(this.PShapeTreesChild, this);

        public override SCSlideMaster ParentSlideMaster => (SCSlideMaster)this.ParentSlide.ParentSlideLayout.ParentSlideMaster;

        public SCPresentation ParentPresentation => this.ParentSlide.ParentPresentation;

        internal SCSlide ParentSlide { get; }
    }
}
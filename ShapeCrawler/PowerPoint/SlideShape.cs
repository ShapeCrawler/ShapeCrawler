using DocumentFormat.OpenXml;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler
{
    internal abstract class SlideShape : Shape
    {
        protected SlideShape(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide, Shape? groupShape)
            : base(pShapeTreeChild, slide, groupShape)
        {
            this.Slide = slide;
        }

        protected SlideShape(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide)
            : base(pShapeTreeChild, slide)
        {
            this.Slide = slide;
        }

        public override IPlaceholder? Placeholder => SlidePlaceholder.Create(this.PShapeTreesChild, this);
        
        public override SCPresentation PresentationInternal => this.Slide.PresentationInternal;

        public ISlide ParentSlide => this.Slide;
        
        internal override SCSlideMaster SlideMasterInternal
        {
            get => (SCSlideMaster)this.Slide.SlideLayout.SlideMaster;

            set
            {
            }
        }

        internal SCSlide Slide { get; }
    }
}
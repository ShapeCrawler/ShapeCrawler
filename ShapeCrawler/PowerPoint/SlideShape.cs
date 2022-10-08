using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;
using OneOf;

namespace ShapeCrawler
{
    internal abstract class SlideShape : Shape
    {
        protected SlideShape(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide, Shape? groupShape)
            : base(pShapeTreeChild, oneOfSlide, groupShape)
        {
            this.Slide = oneOfSlide.Match(slide => slide as SlideBase, layout => layout, master => master);
        }

        public override IPlaceholder? Placeholder => SlidePlaceholder.Create(this.PShapeTreesChild, this);
        
        public override SCPresentation PresentationInternal => this.Slide.PresentationInternal;

        public ISlide ParentSlide => (SCSlide)this.Slide;
        
        internal SlideBase Slide { get; }
    }
}
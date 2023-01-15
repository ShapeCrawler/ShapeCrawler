using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler;

internal abstract class SlideSCShape : SCShape
{
    protected SlideSCShape(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide, SCShape? groupSCShape)
        : base(pShapeTreeChild, oneOfSlide, groupSCShape)
    {
        this.Slide = oneOfSlide.Match(slide => slide as SlideObject, layout => layout, master => master);
    }
    
    protected SlideSCShape(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide)
        : base(pShapeTreeChild, oneOfSlide)
    {
        this.Slide = oneOfSlide.Match(slide => slide as SlideObject, layout => layout, master => master);
    }

    public override IPlaceholder? Placeholder => SlideSCPlaceholder.Create(this.PShapeTreesChild, this);

    internal SlideObject Slide { get; }
}
using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler;

internal abstract class SlideShape : Shape
{
    protected SlideShape(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide, Shape? groupShape)
        : base(pShapeTreeChild, oneOfSlide, groupShape)
    {
        this.Slide = oneOfSlide.Match(slide => slide as SlideObject, layout => layout, master => master);
    }

    public override IPlaceholder? Placeholder => SlidePlaceholder.Create(this.PShapeTreesChild, this);

    internal SlideObject Slide { get; }
}
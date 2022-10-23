using DocumentFormat.OpenXml;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler;

internal abstract class LayoutShape : Shape
{
    protected LayoutShape(SCSlideLayout slideLayout, OpenXmlCompositeElement pShapeTreeChild)
        : base(pShapeTreeChild, slideLayout, null)
    {
        this.SlideLayoutInternal = slideLayout;
    }

    public override IPlaceholder? Placeholder => LayoutPlaceholder.Create(this.PShapeTreesChild, this);

    public SCSlideLayout SlideLayoutInternal { get; }
}
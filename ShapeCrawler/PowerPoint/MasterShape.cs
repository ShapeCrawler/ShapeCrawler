using DocumentFormat.OpenXml;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler;

internal abstract class MasterShape : Shape
{
    protected MasterShape(OpenXmlCompositeElement pShapeTreesChild, SCSlideMaster slideMaster)
        : base(pShapeTreesChild, slideMaster, null)
    {
    }

    public override IPlaceholder Placeholder => MasterPlaceholder.Create(this.PShapeTreesChild);
}
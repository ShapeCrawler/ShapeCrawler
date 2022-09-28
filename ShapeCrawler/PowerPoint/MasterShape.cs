using DocumentFormat.OpenXml;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler;

internal abstract class MasterShape : Shape
{
    protected MasterShape(OpenXmlCompositeElement pShapeTreesChild, SCSlideMaster slideMaster)
        : base(pShapeTreesChild, slideMaster, null)
    {
        this.SlideMasterInternal = slideMaster;
    }

    public override IPlaceholder Placeholder => MasterPlaceholder.Create(this.PShapeTreesChild);

    internal override SCSlideMaster SlideMasterInternal { get; set; }
}
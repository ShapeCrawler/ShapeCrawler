using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tables;

internal class MasterTable : MasterShape, IShape
{
    internal MasterTable(SCSlideMaster slideMasterInternal, P.GraphicFrame pGraphicFrame)
        : base(pGraphicFrame, slideMasterInternal)
    {
    }

    public override SCShapeType ShapeType => SCShapeType.Table;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tables;

internal class LayoutTable : LayoutShape, IShape
{
    internal LayoutTable(SCSlideLayout slideLayoutInternal, P.GraphicFrame pGraphicFrame)
        : base(slideLayoutInternal, pGraphicFrame)
    {
    }

    public override SCShapeType ShapeType => SCShapeType.Table;
}
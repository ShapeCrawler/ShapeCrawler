using DocumentFormat.OpenXml;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class TableShape:Shape
{
    internal TableShape(Position position, ShapeSize shapeSize, ShapeId shapeId, P.GraphicFrame pGraphicFrame)
        :base(position, shapeSize, shapeId, pGraphicFrame)
    {
        var aTable = pGraphicFrame.GetFirstChild<A.Graphic>()!.GetFirstChild<A.GraphicData>()!
            .GetFirstChild<A.Table>() !;
        this.Table = new Table(new TableRowCollection(pGraphicFrame),
            new TableColumnCollection(pGraphicFrame),
            new TableStyleOptions(aTable.TableProperties!),
            pGraphicFrame);
    }
    
    public override ITable? Table { get; }
}
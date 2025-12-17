using DocumentFormat.OpenXml;
using Position = ShapeCrawler.Positions.Position;

// ReSharper disable InconsistentNaming
namespace ShapeCrawler.Shapes;

internal sealed class OleObjectShape(Position position, ShapeSize shapeSize, ShapeId shapeId, OpenXmlElement pShapeTreeElement) 
    : DrawingShape(position, shapeSize, shapeId, pShapeTreeElement);
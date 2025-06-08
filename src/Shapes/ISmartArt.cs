using DocumentFormat.OpenXml;
using ShapeCrawler.Positions;

namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents a SmartArt graphic.
/// </summary>
public interface ISmartArt : IShape
{
    /// <summary>
    ///     Gets the collection of nodes in the SmartArt graphic.
    /// </summary>
    ISmartArtNodeCollection Nodes { get; }
}

internal class SmartArt : Shape, ISmartArt
{
    internal SmartArt(Position position, ShapeSize shapeSize, ShapeId shapeId, OpenXmlElement pShapeTreeElement,
        SmartArtNodeCollection nodeCollection): base(position, shapeSize, shapeId, pShapeTreeElement)
    {
        this.Nodes = nodeCollection;
    }
    public ISmartArtNodeCollection Nodes { get; }
}
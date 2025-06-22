using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Groups;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a group shape.
/// </summary>
public interface IGroup
{
    /// <summary>
    ///     Gets grouped shape collection.
    /// </summary>
    IShapeCollection Shapes { get; }

    /// <summary>
    ///     Gets a shape by name.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T Shape<T>(string groupedShapeName);

    /// <summary>
    ///     Gets a shape by name.
    /// </summary>
    IShape Shape(string groupedShapeName);
}

internal sealed class GroupShape : Shape
{
    private readonly P.GroupShape pGroupShape;

    internal GroupShape(P.GroupShape pGroupShape): base(new Position(pGroupShape), new ShapeSize(pGroupShape), new ShapeId(pGroupShape), pGroupShape)
    {
        this.pGroupShape = pGroupShape;
        this.Shapes = new GroupedShapeCollection(pGroupShape.Elements<OpenXmlCompositeElement>());
        var pShapeProperties = pGroupShape.Descendants<P.ShapeProperties>().First();
    }

    public IShapeCollection Shapes { get; }

    public override Geometry GeometryType => Geometry.Rectangle;

    public bool HasOutline => true;

    public bool HasFill => true;
    
    public IShape Shape(string groupedShapeName) => this.Shapes.Shape(groupedShapeName);

    public T Shape<T>(string groupedShapeName) =>
        (T)this.Shapes.First(groupedShape => groupedShape is T && groupedShape.Name == groupedShapeName);
}
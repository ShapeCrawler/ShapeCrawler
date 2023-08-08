using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents collection of grouped shapes.
/// </summary>
public interface IGroupedShapeCollection : IEnumerable<IShape>
{
    /// <summary>
    ///     Get shape by identifier.
    /// </summary>
    /// <typeparam name="T">The type of shape.</typeparam>
    T GetById<T>(int shapeId)
        where T : IShape;

    /// <summary>
    ///     Get shape by name.
    /// </summary>
    /// <typeparam name="T">The type of shape.</typeparam>
    T GetByName<T>(string shapeName);
}

internal sealed class GroupedShapeCollection : IReadOnlyCollection<IShape>, IGroupedShapeCollection
{
    private readonly List<IShape?> collectionItems;

    internal GroupedShapeCollection(
        P.GroupShape pGroupShapeParam,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        SCGroupShape groupShape,
        TypedOpenXmlPart slideTypedOpenXmlPart,
        List<ImagePart> imageParts)
    {
        var groupedShapes = new List<IShape?>();
        foreach (var child in pGroupShapeParam.ChildElements.OfType<OpenXmlCompositeElement>())
        {
            IShape? shape = null;
            if (child is P.GroupShape pGroupShape)
            {
                shape = new SCGroupShape(pGroupShape, slideOf, groupShape, slideTypedOpenXmlPart, imageParts);
            }
            else if (child is P.Shape pShape)
            {
                var autoShape = new SCAutoShape(pShape, slideOf, groupShape, slideTypedOpenXmlPart);
                autoShape.XChanged += groupShape.OnGroupedShapeXChanged;
                autoShape.YChanged += groupShape.OnGroupedShapeYChanged;
                shape = autoShape;
            }

            if (shape != null)
            {
                groupedShapes.Add(shape);
            }
        }

        this.collectionItems = groupedShapes;
    }

    public int Count => this.collectionItems.Count;

    public T GetById<T>(int shapeId)
        where T : IShape
    {
        var shape = this.collectionItems.First(shape => shape.Id == shapeId);
        return (T)shape;
    }

    public T GetByName<T>(string shapeName)
    {
        var shape = this.collectionItems.First(shape => shape.Name == shapeName);
        return (T)shape;
    }

    public IEnumerator<IShape?> GetEnumerator()
    {
        return this.collectionItems.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
}
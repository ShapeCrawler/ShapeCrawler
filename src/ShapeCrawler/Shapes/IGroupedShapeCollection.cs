using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents collection of grouped shapes.
/// </summary>
public interface IGroupedShapeCollection : IReadOnlyCollection<IShape>
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

internal sealed class SlideGroupedShapes : IGroupedShapeCollection
{
    private readonly List<IShape?> collectionItems;

    internal SlideGroupedShapes(
        P.GroupShape pGroupShape,
        SCSlide slide,
        SCSlideGroupShape groupShape,
        SlidePart sdkSlidePart,
        List<ImagePart> imageParts)
    {
        var groupedShapes = new List<IShape?>();
        foreach (var child in pGroupShape.ChildElements.OfType<OpenXmlCompositeElement>())
        {
            IShape? shape = null;
            if (child is P.GroupShape pGroupShapeItem)
            {
                shape = new SCSlideGroupShape(pGroupShapeItem, slide, groupShape, sdkSlidePart, imageParts);
            }
            else if (child is P.Shape pShape)
            {
                var autoShape = new SCAutoShape(pShape, slide, groupShape, sdkSlidePart);
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
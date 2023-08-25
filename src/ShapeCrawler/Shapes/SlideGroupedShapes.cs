using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class SlideGroupedShapes : IReadOnlyShapeCollection
{
    private readonly SlidePart sdkSlidePart;
    private readonly IEnumerable<OpenXmlCompositeElement> pGroupElements;
    private readonly Lazy<List<IShape>> groupedShapes;

    internal SlideGroupedShapes(SlidePart sdkSlidePart, IEnumerable<OpenXmlCompositeElement> pGroupElements)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.pGroupElements = pGroupElements;
        this.groupedShapes = new Lazy<List<IShape>>(this.ParseGroupedShapes);
    }

    private List<IShape> ParseGroupedShapes()
    {
        var groupedShapes = new List<IShape>();
        foreach (var pGroupElement in pGroupElements)
        {
            IShape? shape = null;
            if (pGroupElement is P.GroupShape pGroupShape)
            {
                shape = new SlideGroupShape(pGroupShape);
            }
            else if (pGroupElement is P.Shape pShape)
            {
                var slideAutoShape = new SlideAutoShape(this.sdkSlidePart, pShape); 
                var groupedAutoShape = new GroupedSlideAutoShape(
                    slideAutoShape,
                    groupShape.OnGroupedShapeXChanged,
                    groupShape.OnGroupedShapeYChanged
                );

                shape = groupedAutoShape;
            }

            if (shape != null)
            {
                groupedShapes.Add(shape);
            }
        }

        return groupedShapes;
    }

    public int Count => this.groupedShapes.Count;

    public T GetById<T>(int shapeId) where T : IShape
    {
        var shape = this.groupedShapes.First(shape => shape.Id == shapeId);
        return (T)shape;
    }

    T IReadOnlyShapeCollection.GetByName<T>(string shapeName)
    {
        throw new System.NotImplementedException();
    }

    public IShape GetByName(string shapeName)
    {
        return this.groupedShapes.First(shape => shape.Name == shapeName);
    }

    public T GetByName<T>(string shapeName)
    {
        var shape = this.groupedShapes.First(shape => shape.Name == shapeName);
        return (T)shape;
    }

    public IEnumerator<IShape> GetEnumerator()
    {
        return this.groupedShapes.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
}
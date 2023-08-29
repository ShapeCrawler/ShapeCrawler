using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideShape;

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
        foreach (var pGroupShapeElement in this.pGroupElements)
        {
            IShape? shape = null;
            if (pGroupShapeElement is P.GroupShape pGroupShape)
            {
                shape = new SlideGroupShape(this.sdkSlidePart, pGroupShape);
            }
            else if (pGroupShapeElement is P.Shape pShape)
            {
                var slideAutoShape = new SlideShape(this.sdkSlidePart, pShape); 
                var groupedAutoShape = new GroupedSlideShape(slideAutoShape);

                shape = groupedAutoShape;
            }

            if (shape != null)
            {
                groupedShapes.Add(shape);
            }
        }

        return groupedShapes;
    }

    public int Count => this.groupedShapes.Value.Count;

    public T GetById<T>(int shapeId) where T : IShape
    {
        var shape = this.groupedShapes.Value.First(shape => shape.Id == shapeId);
        return (T)shape;
    }

    T IReadOnlyShapeCollection.GetByName<T>(string shapeName)
    {
        throw new System.NotImplementedException();
    }

    public IShape GetByName(string shapeName)
    {
        return this.groupedShapes.Value.First(shape => shape.Name == shapeName);
    }

    public T GetByName<T>(string shapeName)
    {
        var shape = this.groupedShapes.Value.First(shape => shape.Name == shapeName);
        return (T)shape;
    }

    public IEnumerator<IShape> GetEnumerator()
    {
        return this.groupedShapes.Value.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    public IShape this[int index] => throw new NotImplementedException();
}
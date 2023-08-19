using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class SCSlideGroupedShapes : IReadOnlyShapeCollection
{
    private readonly List<IShape> groupedShapes;

    internal SCSlideGroupedShapes(P.GroupShape parentPGroupShape, SCSlideGroupShape groupShape)
    {
        var groupedShapes = new List<IShape?>();
        foreach (var parentPGroupShapeChild in parentPGroupShape.ChildElements.OfType<OpenXmlCompositeElement>())
        {
            IShape? shape = null;
            if (parentPGroupShapeChild is P.GroupShape pGroupShape)
            {
                shape = new SCSlideGroupShape(pGroupShape, this, sdkSlidePart, imageParts);
            }
            else if (parentPGroupShapeChild is P.Shape pShape)
            {
                // var autoShape = new SCSlideAutoShape(pShape, groupShape, sdkSlidePart, groupShape.OnGroupedShapeXChanged, groupShape.OnGroupedShapeYChanged);
                var slideGroupedAutoShape = new SCSlideGroupedAutoShape(
                    new SlideAutoShape(pShape, this, sdkSlidePart),
                    groupShape.OnGroupedShapeXChanged, groupShape.OnGroupedShapeYChanged);

                shape = slideGroupedAutoShape;
            }

            if (shape != null)
            {
                groupedShapes.Add(shape);
            }
        }

        this.groupedShapes = groupedShapes;
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
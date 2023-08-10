using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Shapes;

namespace ShapeCrawler;

internal sealed class SCSlideGroupedShapeCollection : IReadOnlyShapeCollection
{
    private readonly List<IShape?> collectionItems;

    internal SCSlideGroupedShapeCollection(
        DocumentFormat.OpenXml.Presentation.GroupShape parentPGroupShape,
        SCSlideGroupShape groupShape,
        SlidePart sdkSlidePart,
        List<ImagePart> imageParts)
    {
        var groupedShapes = new List<IShape?>();
        foreach (var parentPGroupShapeChild in parentPGroupShape.ChildElements.OfType<OpenXmlCompositeElement>())
        {
            IShape? shape = null;
            if (parentPGroupShapeChild is DocumentFormat.OpenXml.Presentation.GroupShape pGroupShape)
            {
                shape = new SCSlideGroupShape(pGroupShape, groupShape, sdkSlidePart, imageParts);
            }
            else if (parentPGroupShapeChild is DocumentFormat.OpenXml.Presentation.Shape pShape)
            {
                // var autoShape = new SCSlideAutoShape(pShape, groupShape, sdkSlidePart, groupShape.OnGroupedShapeXChanged, groupShape.OnGroupedShapeYChanged);
                var slideGroupedAutoShape = new SCSlideGroupedAutoShape(
                    new SCSlideAutoShape(pShape, this, sdkSlidePart),
                    groupShape.OnGroupedShapeXChanged, groupShape.OnGroupedShapeYChanged);

                shape = slideGroupedAutoShape;
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
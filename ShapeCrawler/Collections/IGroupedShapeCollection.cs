using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Factories;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections;

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

internal sealed class GroupedShapeCollection : LibraryCollection<IShape>, IGroupedShapeCollection
{
    private GroupedShapeCollection(List<IShape> groupedShapes)
        : base(groupedShapes)
    {
    }

    public T GetById<T>(int shapeId)
        where T : IShape
    {
        var shape = this.CollectionItems.First(shape => shape.Id == shapeId);
        return (T)shape;
    }

    public T GetByName<T>(string shapeName)
    {
        var shape = this.CollectionItems.First(shape => shape.Name == shapeName);
        return (T)shape;
    }

    internal static GroupedShapeCollection Create(P.GroupShape pGroupShapeParam, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide, SCGroupShape groupShape)
    {
        var autoShapeCreator = new AutoShapeCreator();
        var oleGrFrameHandler = new OleGraphicFrameHandler();
        var pictureHandler = new PictureHandler();
        var chartGrFrameHandler = new ChartGraphicFrameHandler();
        var tableGrFrameHandler = new TableGraphicFrameHandler();

        autoShapeCreator.Successor = oleGrFrameHandler;
        oleGrFrameHandler.Successor = pictureHandler;
        pictureHandler.Successor = chartGrFrameHandler;
        chartGrFrameHandler.Successor = tableGrFrameHandler;

        var groupedShapes = new List<IShape>();
        foreach (var child in pGroupShapeParam.ChildElements.OfType<OpenXmlCompositeElement>())
        {
            Shape? shape;
            if (child is P.GroupShape pGroupShape)
            {
                shape = new SCGroupShape(pGroupShape, oneOfSlide, groupShape);
            }
            else
            {
                shape = autoShapeCreator.Create(child, oneOfSlide, groupShape);
            }

            if (shape != null)
            {
                groupedShapes.Add(shape);
            }
        }

        return new GroupedShapeCollection(groupedShapes);
    }
}
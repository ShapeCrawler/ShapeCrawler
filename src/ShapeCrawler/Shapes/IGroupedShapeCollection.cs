using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Services;
using ShapeCrawler.Services.Factories;
using ShapeCrawler.Shapes;
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
    private readonly List<IShape> collectionItems;

    internal GroupedShapeCollection (
        P.GroupShape pGroupShapeParam,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        SCGroupShape groupShape,
        TypedOpenXmlPart slideTypedOpenXmlPart,
        List<ImagePart> imageParts) 
    {
        var autoShapeCreator = new AutoShapeCreator();
        var oleGrFrameHandler = new OleGraphicFrameHandler();
        var pictureHandler = new PictureHandler(imageParts, slideTypedOpenXmlPart);
        var chartGrFrameHandler = new ChartGraphicFrameHandler();
        var tableGrFrameHandler = new TableGraphicFrameHandler();

        autoShapeCreator.Successor = oleGrFrameHandler;
        oleGrFrameHandler.Successor = pictureHandler;
        pictureHandler.Successor = chartGrFrameHandler;
        chartGrFrameHandler.Successor = tableGrFrameHandler;

        var groupedShapes = new List<IShape>();
        foreach (var child in pGroupShapeParam.ChildElements.OfType<OpenXmlCompositeElement>())
        {
            SCShape? shape;
            if (child is P.GroupShape pGroupShape)
            {
                shape = new SCGroupShape(pGroupShape, slideOf, groupShape, slideTypedOpenXmlPart, imageParts);
            }
            else
            {
                shape = autoShapeCreator.FromTreeChild(child, slideOf, groupShape, slideTypedOpenXmlPart);
                if (shape != null)
                {
                    shape.XChanged += groupShape.OnGroupedShapeXChanged;    
                    shape.YChanged += groupShape.OnGroupedShapeYChanged;    
                }
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

    public IEnumerator<IShape> GetEnumerator()
    {
        return this.collectionItems.GetEnumerator();
    }
    
    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
}
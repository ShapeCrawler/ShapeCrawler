using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Assets;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

/// <summary>
///     Represents placeholder shapes collection.
/// </summary>
internal sealed class PlaceholderShapes(IUserSlideShapeCollection shapes, SlidePart slidePart)
{
    /// <summary>
    ///     Adds a date and time placeholder.
    /// </summary>
    internal IShape AddDateAndTime()
    {
        // Check if a DateAndTime placeholder already exists
        var existingDateTimePlaceholder = shapes
            .FirstOrDefault(shape => shape.PlaceholderType == PlaceholderType.DateAndTime);

        if (existingDateTimePlaceholder != null)
        {
            throw new SCException("The slide already contains a Date and Time placeholder.");
        }

        // Load the date-time placeholder XML template
        var xml = new AssetCollection(Assembly.GetExecutingAssembly()).StringOf("date and time placeholder.xml");
        var pShape = new P.Shape(xml);
        var nextShapeId = new NewShapeProperties(shapes).Id();

        // Append the shape to the slide
        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pShape);

        // Get the added shape and set its properties
        var addedShape = shapes.Last<TextShape>();
        addedShape.Id = nextShapeId;
        addedShape.Name = $"Date Placeholder {nextShapeId}";

        return addedShape;
    }

    /// <summary>
    ///     Adds a footer placeholder.
    /// </summary>
    internal IShape AddFooter()
    {
        // Check if a Footer placeholder already exists
        var existingFooterPlaceholder = shapes
            .FirstOrDefault(shape => shape.PlaceholderType == PlaceholderType.Footer);

        if (existingFooterPlaceholder != null)
        {
            throw new SCException("The slide already contains a Footer placeholder.");
        }

        // Load the footer placeholder XML template
        var xml = new AssetCollection(Assembly.GetExecutingAssembly()).StringOf("footer placeholder.xml");
        var pShape = new P.Shape(xml);
        var nextShapeId = new NewShapeProperties(shapes).Id();

        // Append the shape to the slide
        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pShape);

        // Get the added shape and set its properties
        var addedShape = shapes.Last<TextShape>();
        addedShape.Id = nextShapeId;
        addedShape.Name = $"Footer Placeholder {nextShapeId}";

        return addedShape;
    }

    /// <summary>
    ///     Adds a slide number placeholder.
    /// </summary>
    internal IShape AddSlideNumber()
    {
        // Check if a Slide Number placeholder already exists
        var existingSlideNumberPlaceholder = shapes
            .FirstOrDefault(shape => shape.PlaceholderType == PlaceholderType.SlideNumber);

        if (existingSlideNumberPlaceholder != null)
        {
            throw new SCException("The slide already contains a Slide Number placeholder.");
        }

        // Load the slide number placeholder XML template
        var xml = new AssetCollection(Assembly.GetExecutingAssembly()).StringOf("slide number placeholder.xml");
        var pShape = new P.Shape(xml);
        var nextShapeId = new NewShapeProperties(shapes).Id();

        // Append the shape to the slide
        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pShape);

        // Get the added shape and set its properties
        var addedShape = shapes.Last<TextShape>();
        addedShape.Id = nextShapeId;
        addedShape.Name = $"Slide Number Placeholder {nextShapeId}";

        return addedShape;
    }
}

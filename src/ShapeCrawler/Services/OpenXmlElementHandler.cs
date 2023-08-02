using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.Services.Factories;

internal abstract class OpenXmlElementHandler
{
    internal OpenXmlElementHandler? Successor { get; set; }

    internal abstract SCShape? FromTreeChild(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        OneOf<ShapeCollection, SCGroupShape> shapeCollectionOf);
}
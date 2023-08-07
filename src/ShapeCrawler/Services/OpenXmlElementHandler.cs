using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.Services;

internal abstract class OpenXmlElementHandler
{
    internal OpenXmlElementHandler? Successor { get; set; }

    internal abstract SCShape? FromTreeChild(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        OneOf<ShapeCollection, SCGroupShape> shapeCollectionOf,
        TypedOpenXmlPart slideTypedOpenXmlPart);
}
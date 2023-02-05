using OneOf;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler.Factories;

internal abstract class OpenXmlElementHandler
{
    internal OpenXmlElementHandler? Successor { get; set; }

    internal abstract SCShape? Create(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject,
        OneOf<ShapeCollection, SCGroupShape> shapeCollection);
}
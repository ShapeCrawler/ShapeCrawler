using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler.Factories;

internal abstract class OpenXmlElementHandler
{
    internal OpenXmlElementHandler? Successor { get; set; }

    internal abstract Shape? Create(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject, SCGroupShape groupShape);
}
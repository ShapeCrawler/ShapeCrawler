using DocumentFormat.OpenXml;
using ShapeCrawler.Collections;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders;

internal class SlidePlaceholder : Placeholder
{
    private readonly SlideShape slideShape;

    private SlidePlaceholder(P.PlaceholderShape pPlaceholderShape, SlideShape slideShape)
        : base(pPlaceholderShape)
    {
        this.slideShape = slideShape;
        this.ReferencedShapeLazy = new ResettableLazy<Shape>(this.GetReferencedShape);
    }

    internal static SlidePlaceholder? Create(OpenXmlCompositeElement pShapeTreeChild, SlideShape slideShape)
    {
        var pPlaceholder = pShapeTreeChild.GetPNvPr().GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholder == null)
        {
            return null;
        }

        return new SlidePlaceholder(pPlaceholder, slideShape);
    }

    private Shape? GetReferencedShape()
    {
        if (this.slideShape.SlideBase is SCSlideLayout slideLayout)
        {
            var masterShapes = slideLayout.SlideMasterInternal.ShapesInternal;
            return masterShapes.GetReferencedShapeOrDefault(this.PPlaceholderShape);
        }

        if (this.slideShape.SlideBase is SCSlideMaster)
        {
            return null;
        }

        var slide = (SCSlide)this.slideShape.SlideBase;
        var layout = (SCSlideLayout)slide.SlideLayout;
        var layoutShapes = layout.ShapesInternal;
        var referencedShape = layoutShapes.GetReferencedShapeOrDefault(this.PPlaceholderShape);

        if (referencedShape == null)
        {
            var masterShapes = layout.SlideMasterInternal.ShapesInternal;
            return masterShapes.GetReferencedShapeOrDefault(this.PPlaceholderShape);
        }

        return referencedShape;
    }
}
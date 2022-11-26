using DocumentFormat.OpenXml;
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
    }

    internal override ResettableLazy<Shape?> ReferencedShape => new (this.GetReferencedShape);

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
            return masterShapes.GetReferencedShapeOrNull(this.PPlaceholderShape);
        }

        if (this.slideShape.SlideBase is SCSlideMaster)
        {
            return null;
        }

        var slide = (SCSlide)this.slideShape.SlideBase;
        var layout = (SCSlideLayout)slide.SlideLayout;
        var layoutShapes = layout.ShapesInternal;
        var referencedShape = layoutShapes.GetReferencedShapeOrNull(this.PPlaceholderShape);

        if (referencedShape == null)
        {
            var masterShapes = layout.SlideMasterInternal.ShapesInternal;
            return masterShapes.GetReferencedShapeOrNull(this.PPlaceholderShape);
        }

        return referencedShape;
    }
}
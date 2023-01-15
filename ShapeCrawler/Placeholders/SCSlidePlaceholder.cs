using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders;

internal sealed class SCSlidePlaceholder : SCPlaceholder
{
    private readonly SCSlideShape _slideSCShape;

    private SCSlidePlaceholder(P.PlaceholderShape pPlaceholderShape, SCSlideShape slideSCShape)
        : base(pPlaceholderShape)
    {
        this._slideSCShape = slideSCShape;
    }

    internal override ResettableLazy<SCShape?> ReferencedShape => new (this.GetReferencedShape);

    internal static SCSlidePlaceholder? Create(OpenXmlCompositeElement pShapeTreeChild, SCSlideShape slideSCShape)
    {
        var pPlaceholder = pShapeTreeChild.GetPNvPr().GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholder == null)
        {
            return null;
        }

        return new SCSlidePlaceholder(pPlaceholder, slideSCShape);
    }

    private SCShape? GetReferencedShape()
    {
        if (this._slideSCShape.SlideBase is SCSlideLayout slideLayout)
        {
            var masterShapes = slideLayout.SlideMasterInternal.ShapesInternal;
            return masterShapes.GetReferencedShapeOrNull(this.PPlaceholderShape);
        }

        if (this._slideSCShape.SlideBase is SCSlideMaster)
        {
            return null;
        }

        var slide = (SCSlide)this._slideSCShape.SlideBase;
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
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders;

internal sealed class SlideSCPlaceholder : SCPlaceholder
{
    private readonly SlideSCShape _slideSCShape;

    private SlideSCPlaceholder(P.PlaceholderShape pPlaceholderShape, SlideSCShape slideSCShape)
        : base(pPlaceholderShape)
    {
        this._slideSCShape = slideSCShape;
    }

    internal override ResettableLazy<SCShape?> ReferencedShape => new (this.GetReferencedShape);

    internal static SlideSCPlaceholder? Create(OpenXmlCompositeElement pShapeTreeChild, SlideSCShape slideSCShape)
    {
        var pPlaceholder = pShapeTreeChild.GetPNvPr().GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholder == null)
        {
            return null;
        }

        return new SlideSCPlaceholder(pPlaceholder, slideSCShape);
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
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
        this.referencedShape = new ResettableLazy<Shape>(this.GetReferencedShape);
    }
    
    internal static SlidePlaceholder? Create(OpenXmlCompositeElement pShapeTreeChild, SlideShape slideShape)
    {
        var pPlaceholder = pShapeTreeChild.ApplicationNonVisualDrawingProperties().GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholder == null)
        {
            return null;
        }

        return new SlidePlaceholder(pPlaceholder, slideShape);
    }

    private Shape GetReferencedShape()
    {
        if (this.slideShape.SlideBase is SCSlideLayout slideLayout)
        {
            var shapes = slideLayout.SlideMasterInternal.ShapesInternal;
            return shapes.GetReferencedShapeOrDefault(this.PPlaceholderShape);
        }
        
        if (this.slideShape.SlideBase is SCSlideMaster slideMaster)
        {
            return null;
        }
        
        var layout = (SCSlideLayout)this.slideShape.ParentSlide.SlideLayout;
        var layoutShapes = layout.ShapesInternal;
        var referencedShape = layoutShapes.GetReferencedShapeOrDefault(this.PPlaceholderShape);

        return referencedShape;
    }
}
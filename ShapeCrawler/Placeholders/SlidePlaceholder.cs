using DocumentFormat.OpenXml;
using ShapeCrawler.Collections;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders
{
    /// <summary>
    ///     Represents a placeholder located on a slide.
    /// </summary>
    internal class SlidePlaceholder : Placeholder
    {
        private readonly SlideShape slideShape;

        private SlidePlaceholder(P.PlaceholderShape pPlaceholderShape, SlideShape slideShape)
            : base(pPlaceholderShape)
        {
            this.slideShape = slideShape;
            this.layoutReferencedShape = new ResettableLazy<Shape>(this.GetReferencedShape);
        }

        /// <summary>
        ///     Creates placeholder. Returns <c>NULL</c> if the specified shape is not placeholder.
        /// </summary>
        public static SlidePlaceholder? Create(OpenXmlCompositeElement pShapeTreeChild, SlideShape slideShape)
        {
            var pPlaceholderShape =
                pShapeTreeChild.ApplicationNonVisualDrawingProperties().GetFirstChild<P.PlaceholderShape>();
            if (pPlaceholderShape == null)
            {
                return null;
            }

            return new SlidePlaceholder(pPlaceholderShape, slideShape);
        }

        private Shape GetReferencedShape()
        {
            var layout = this.slideShape.Slide.SlideLayoutInternal;
            var layoutShapes = layout.ShapesInternal;
            var referencedShape = layoutShapes.GetReferencedShapeOrDefault(this.PPlaceholderShape);

            if (referencedShape != null)
            {
                return referencedShape;
            }

            var masterShapes = (ShapeCollection)layout.SlideMasterInternal.Shapes;

            return masterShapes.GetReferencedShapeOrDefault(this.PPlaceholderShape);
        }
    }
}
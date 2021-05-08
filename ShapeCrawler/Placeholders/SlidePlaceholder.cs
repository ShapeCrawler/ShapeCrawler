using DocumentFormat.OpenXml;
using ShapeCrawler.Collections;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders
{
    /// <summary>
    ///     Represents a placeholder located on a slide.
    /// </summary>
    internal class SlidePlaceholder : Placeholder
    {
        private readonly SlideShape parentSlideShape;

        public SlidePlaceholder(P.PlaceholderShape sdkPPlaceholderShape, SlideShape parentSlideShape)
            : base(sdkPPlaceholderShape)
        {
            this.parentSlideShape = parentSlideShape;
            this.layoutReferencedShape = new ResettableLazy<Shape>(() => this.GetReferencedLayoutShape());
        }

        private Shape GetReferencedLayoutShape()
        {
            ShapeCollection shapes = (ShapeCollection)this.parentSlideShape.ParentSlide.ParentSlideLayout.Shapes;

            return shapes.GetShapeByPPlaceholderShape(this.SdkPPlaceholderShape);
        }

        /// <summary>
        ///     Creates placeholder. Returns <c>NULL</c> if the specified shape is not placeholder.
        /// </summary>
        internal static SlidePlaceholder Create(OpenXmlCompositeElement pShapeTreeChild, SlideShape slideShape)
        {
            P.PlaceholderShape pPlaceholderShape =
                pShapeTreeChild.ApplicationNonVisualDrawingProperties().GetFirstChild<P.PlaceholderShape>();
            if (pPlaceholderShape == null)
            {
                return null;
            }

            return new SlidePlaceholder(pPlaceholderShape, slideShape);
        }
    }
}
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
        #region Constructors

        public SlidePlaceholder(P.PlaceholderShape pPlaceholderShape, SlideShape slideShape)
            : base(pPlaceholderShape)
        {
            this.BaseShape = new ResettableLazy<Shape>(() =>
                ((ShapeCollection)slideShape.Slide.ParentSlideLayout.Shapes).GetShapeByPPlaceholderShape(pPlaceholderShape));
        }

        #endregion Constructors

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
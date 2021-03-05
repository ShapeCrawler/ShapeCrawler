using DocumentFormat.OpenXml;
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
        private readonly ResettableLazy<Shape> _layoutShape;
        private SlideAutoShape slideAutoShape;

        #region Constructors

        public SlidePlaceholder(P.PlaceholderShape pPlaceholderShape, SlideShape slideShape)
            : base(pPlaceholderShape)
        {
            _layoutShape = new ResettableLazy<Shape>(() =>
                slideShape.Slide.SlideLayout.Shapes.GetShapeByPPlaceholderShape(pPlaceholderShape));
        }

        public SlidePlaceholder(P.PlaceholderShape pPlaceholderShape, SlideAutoShape slideAutoShape)
            : base(pPlaceholderShape)
        {
            _layoutShape = new ResettableLazy<Shape>(() =>
                slideAutoShape.Slide.SlideLayout.Shapes.GetShapeByPPlaceholderShape(pPlaceholderShape));
            Shape = _layoutShape.Value;
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

        internal static SlidePlaceholder Create(SlideAutoShape slideAutoShape)
        {
            P.PlaceholderShape pPlaceholderShape = slideAutoShape.PShapeTreeChild.ApplicationNonVisualDrawingProperties().GetFirstChild<P.PlaceholderShape>();
            if (pPlaceholderShape == null)
            {
                return null;
            }

            return new SlidePlaceholder(pPlaceholderShape, slideAutoShape);
        }
    }
}
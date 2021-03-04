using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;

namespace ShapeCrawler.Placeholders
{
    internal class LayoutPlaceholder : Placeholder
    {
        private readonly ResettableLazy<Shape> _masterShape;

        public LayoutPlaceholder(DocumentFormat.OpenXml.Presentation.PlaceholderShape pPlaceholderShape, LayoutShape layoutShape)
            : base(pPlaceholderShape)
        {
            _masterShape = new ResettableLazy<Shape>(() =>
                layoutShape.SlideLayout.SlideMaster.Shapes.GetShapeByPPlaceholderShape(pPlaceholderShape));
        }

        /// <summary>
        ///     Creates placeholder. Returns <c>NULL</c> if the specified shape is not placeholder.
        /// </summary>
        internal static LayoutPlaceholder Create(OpenXmlCompositeElement pShapeTreeChild, LayoutShape slideShape)
        {
            DocumentFormat.OpenXml.Presentation.PlaceholderShape pPlaceholderShape =
                pShapeTreeChild.ApplicationNonVisualDrawingProperties().GetFirstChild<DocumentFormat.OpenXml.Presentation.PlaceholderShape>();
            if (pPlaceholderShape == null)
            {
                return null;
            }

            return new LayoutPlaceholder(pPlaceholderShape, slideShape);
        }
    }
}
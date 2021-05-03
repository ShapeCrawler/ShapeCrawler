using DocumentFormat.OpenXml;
using ShapeCrawler.Collections;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders
{
    internal class LayoutPlaceholder : Placeholder
    {
        private LayoutPlaceholder(P.PlaceholderShape pPlaceholderShape, LayoutShape layoutShape)
            : base(pPlaceholderShape)
        {
            BaseShape = new ResettableLazy<Shape>(() =>
                ((ShapeCollection)layoutShape.SlideLayout.ParentSlideMaster.Shapes).GetShapeByPPlaceholderShape(pPlaceholderShape));
        }

        /// <summary>
        ///     Creates placeholder. Returns <c>NULL</c> if the specified shape is not placeholder.
        /// </summary>
        internal static LayoutPlaceholder Create(OpenXmlCompositeElement pShapeTreeChild, LayoutShape slideShape)
        {
            P.PlaceholderShape pPlaceholderShape =
                pShapeTreeChild.ApplicationNonVisualDrawingProperties().GetFirstChild<P.PlaceholderShape>();
            if (pPlaceholderShape == null)
            {
                return null;
            }

            return new LayoutPlaceholder(pPlaceholderShape, slideShape);
        }
    }
}
using DocumentFormat.OpenXml;
using ShapeCrawler.Collections;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders
{
    /// <summary>
    ///     Represents a placeholder located on a Slide Layout.
    /// </summary>
    internal class LayoutPlaceholder : Placeholder
    {
        private readonly LayoutShape parentLayoutShape;

        private LayoutPlaceholder(P.PlaceholderShape pPlaceholderShape, LayoutShape parentLayoutShape)
            : base(pPlaceholderShape)
        {
            this.parentLayoutShape = parentLayoutShape;
            this.layoutReferencedShape = new ResettableLazy<Shape>(this.GetReferencedShape);
        }

        internal static LayoutPlaceholder Create(OpenXmlCompositeElement pShapeTreeChild, LayoutShape layoutShape)
        {
            P.PlaceholderShape pPlaceholderShape =
                pShapeTreeChild.ApplicationNonVisualDrawingProperties().GetFirstChild<P.PlaceholderShape>();
            if (pPlaceholderShape == null)
            {
                return null;
            }

            return new LayoutPlaceholder(pPlaceholderShape, layoutShape);
        }

        private Shape GetReferencedShape()
        {
            ShapeCollection shapeCollection = (ShapeCollection)this.parentLayoutShape.ParentSlideLayoutInternal.ParentSlideMaster.Shapes;

            return shapeCollection.GetReferencedShapeOrDefault(this.PPlaceholderShape);
        }
    }
}
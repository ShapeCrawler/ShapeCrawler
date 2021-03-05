using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders
{
    /// <summary>
    ///     Represents a placeholder located on a Slide Master.
    /// </summary>
    internal class MasterPlaceholder : Placeholder
    {
        public MasterPlaceholder(P.PlaceholderShape pPlaceholderShape) : base(pPlaceholderShape)
        {

        }

        internal static MasterPlaceholder Create(MasterAutoShape masterAutoShape)
        {
            P.PlaceholderShape pPlaceholderShape = masterAutoShape.PShapeTreeChild.ApplicationNonVisualDrawingProperties().GetFirstChild<P.PlaceholderShape>();
            if (pPlaceholderShape == null)
            {
                return null;
            }

            return new MasterPlaceholder(pPlaceholderShape);
        }

        /// <summary>
        ///     Creates placeholder. Returns <c>NULL</c> if the specified shape is not placeholder.
        /// </summary>
        internal static MasterPlaceholder Create(OpenXmlCompositeElement pShapeTreeChild)
        {
            P.PlaceholderShape pPlaceholderShape =
                pShapeTreeChild.ApplicationNonVisualDrawingProperties().GetFirstChild<P.PlaceholderShape>();
            if (pPlaceholderShape == null)
            {
                return null;
            }

            return new MasterPlaceholder(pPlaceholderShape);
        }
    }
}
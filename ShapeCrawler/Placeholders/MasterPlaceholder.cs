using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;

namespace ShapeCrawler.Placeholders
{
    /// <summary>
    ///     Represents a placeholder located on a Slide Master.
    /// </summary>
    internal class MasterPlaceholder : Placeholder
    {
        public MasterPlaceholder(DocumentFormat.OpenXml.Presentation.PlaceholderShape pPlaceholderShape) : base(pPlaceholderShape)
        {

        }

        /// <summary>
        ///     Creates placeholder. Returns <c>NULL</c> if the specified shape is not placeholder.
        /// </summary>
        internal static MasterPlaceholder Create(OpenXmlCompositeElement pShapeTreeChild)
        {
            DocumentFormat.OpenXml.Presentation.PlaceholderShape pPlaceholderShape =
                pShapeTreeChild.ApplicationNonVisualDrawingProperties().GetFirstChild<DocumentFormat.OpenXml.Presentation.PlaceholderShape>();
            if (pPlaceholderShape == null)
            {
                return null;
            }

            return new MasterPlaceholder(pPlaceholderShape);
        }
    }
}
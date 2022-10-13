using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders
{
    /// <summary>
    ///     Represents a placeholder located on a Slide Master.
    /// </summary>
    internal class MasterPlaceholder : Placeholder
    {
        public MasterPlaceholder(P.PlaceholderShape pPlaceholderShape)
            : base(pPlaceholderShape)
        {
            // Slide Master is the lowest slide level, therefore its placeholders do not have referenced shape.
            this.ReferencedShapeLazy = new ResettableLazy<Shape?>(() => null);
        }

        internal static MasterPlaceholder? Create(OpenXmlCompositeElement pShapeTreeChild)
        {
            P.PlaceholderShape pPlaceholderShape =
                pShapeTreeChild.GetPNvPr().GetFirstChild<P.PlaceholderShape>();
            if (pPlaceholderShape == null)
            {
                return null;
            }

            return new MasterPlaceholder(pPlaceholderShape);
        }
    }
}
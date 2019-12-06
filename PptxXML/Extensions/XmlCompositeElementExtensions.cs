using System.Linq;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Extensions
{
    /// <summary>
    /// Extension methods for <see cref="OpenXmlCompositeElement"/> instance.
    /// </summary>
    public static class XmlCompositeElementExtensions
    {
        #region Fields

        private const uint TitlePlaceholderIndexValue = 0;

        #endregion

        /// <summary>
        /// Get placeholder index.
        /// </summary>
        /// <param name="xmlCompositeElement">
        /// An element which can be located on slide or master slide.
        /// </param>
        /// <returns></returns>
        public static uint? GetPlaceholderIndex(this OpenXmlCompositeElement xmlCompositeElement)
        {
            var ph = xmlCompositeElement.Descendants<P.PlaceholderShape>().FirstOrDefault();
            if (ph == null)
            {
                return null;
            }

            // Simple title and centered title placeholders were united.
            var phType = ph.Type;
            if (phType != null && (phType == P.PlaceholderValues.Title || phType == P.PlaceholderValues.CenteredTitle))
            {
                return TitlePlaceholderIndexValue;
            }

            return ph.Index.Value;
        }

        /// <summary>
        /// Gets non visual drawing properties values (cNvPr).
        /// </summary>
        /// <returns>(identifier, hidden)</returns>
        public static (int, bool) GetNvPrValues(this OpenXmlCompositeElement compositeElement)
        {
            // .First() is used instead .Single() because group shape can have more than one id for its child elements
            var cNvPr = compositeElement.Descendants<P.NonVisualDrawingProperties>().First();
            var id = (int) cNvPr.Id.Value;
            var parsedHiddenValue = cNvPr.Hidden?.Value;
            var hidden = parsedHiddenValue != null && parsedHiddenValue == true;

            return (id, hidden);
        }

        /// <summary>
        /// Gets identifier.
        /// </summary>
        public static int GetId(this OpenXmlCompositeElement compositeElement)
        {
            // .First() is used instead .Single() because group shape can have more than one id for its child elements
            var cNvPr = compositeElement.Descendants<P.NonVisualDrawingProperties>().First();
            var id = (int)cNvPr.Id.Value;

            return id;
        }

        /// <summary>
        /// Determines whether element is chart. 
        /// </summary>
        /// <param name="compositeElement"></param>        
        public static bool IsChart(this OpenXmlCompositeElement compositeElement)
        {
            var grData = compositeElement.Descendants<D.GraphicData>().SingleOrDefault();
            if (grData == null)
            {
                return false;
            }
            var endsWithChart = grData?.Uri?.Value?.EndsWith("chart");
            return endsWithChart != null && endsWithChart != false;
        }

        /// <summary>
        /// Determines whether element is table. 
        /// </summary>
        /// <param name="compositeElement"></param>        
        public static bool IsTable(this OpenXmlCompositeElement compositeElement)
        {
            return compositeElement.Descendants<D.Table>().Any();
        }
    }
}

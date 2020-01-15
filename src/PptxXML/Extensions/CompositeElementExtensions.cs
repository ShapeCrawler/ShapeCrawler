using System.Linq;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Extensions
{
    /// <summary>
    /// Extension methods for <see cref="OpenXmlCompositeElement"/> instance.
    /// </summary>
    public static class CompositeElementExtensions
    {
        #region Fields

        private const uint TitleIndexValue = 100; // Title and CenteredTitle have same custom index value
        private const uint SubtitleIndexValue = 101;

        #endregion Fields

        /// <summary>
        /// Parses <see cref="P.PlaceholderValues.Title"/> and <see cref="P.PlaceholderValues.CenteredTitle"/> placeholders.
        /// </summary>
        /// <param name="xmlCompositeElement">
        /// An element which can be located on slide or master slide.
        /// </param>
        /// <returns>Placeholder index or null.</returns>
        public static uint? GetPlaceholderIndex(this OpenXmlCompositeElement xmlCompositeElement)
        {
            var ph = xmlCompositeElement.Descendants<P.PlaceholderShape>().FirstOrDefault();
            if (ph == null)
            {
                return null;
            }

            // Simple title and centered title placeholders were united.
            var phType = ph.Type;
            if (phType != null)
            {
                if (phType == P.PlaceholderValues.Title || phType == P.PlaceholderValues.CenteredTitle)
                {
                    return TitleIndexValue;
                }
                if (phType == P.PlaceholderValues.SubTitle)
                {
                    return SubtitleIndexValue;
                }
            }

            var index = ph.Index;

            if (index == null)
            {
                return null;
            }

            return index.Value;
        }

        /// <summary>
        /// Determines whether element is placeholder.
        /// </summary>
        /// <param name="xmlCompositeElement"></param>
        /// <returns></returns>
        public static bool IsPlaceholder(this OpenXmlCompositeElement xmlCompositeElement)
        {
            return xmlCompositeElement.Descendants<P.PlaceholderShape>().Any();
        }

        /// <summary>
        /// Gets non visual drawing properties values (cNvPr).
        /// </summary>
        /// <returns>(identifier, hidden, name)</returns>
        public static (int, bool, string) GetNvPrValues(this OpenXmlCompositeElement compositeElement)
        {
            // .First() is used instead .Single() because group shape can have more than one id for its child elements
            var cNvPr = compositeElement.Descendants<P.NonVisualDrawingProperties>().First();
            var id = (int) cNvPr.Id.Value;
            var name = cNvPr.Name.Value;
            var parsedHiddenValue = cNvPr.Hidden?.Value;
            var hidden = parsedHiddenValue != null && parsedHiddenValue == true;

            return (id, hidden, name);
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

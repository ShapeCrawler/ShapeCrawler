using DocumentFormat.OpenXml;
using PptxXML.Enums;

namespace PptxXML.Models.Elements
{
    /// <summary>
    /// Represents a chart.
    /// </summary>
    public class ChartEx: Element
    {
        #region Constructors

        /// <summary>
        /// Initialise an instance of <see cref="ChartEx"/> class.
        /// </summary>
        /// <param name="xmlCompositeElement"></param>
        public ChartEx(OpenXmlCompositeElement xmlCompositeElement) :
            base(xmlCompositeElement)
        {
            Type = ElementType.Chart;
        }

        #endregion
    }
}

using DocumentFormat.OpenXml;
using PptxXML.Entities.Elements;
using PptxXML.Enums;

namespace PptxXML.Entities.Elements
{
    /// <summary>
    /// Represent a chart.
    /// </summary>
    public class Chart: Element
    {
        #region Constructors

        /// <summary>
        /// Initialise an instance of <see cref="Chart"/> class.
        /// </summary>
        /// <param name="xmlCompositeElement"></param>
        public Chart(OpenXmlCompositeElement xmlCompositeElement) :
            base(xmlCompositeElement)
        {
            Type = ElementType.Chart;
        }

        #endregion
    }
}

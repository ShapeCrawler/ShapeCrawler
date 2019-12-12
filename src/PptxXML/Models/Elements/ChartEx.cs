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
        /// Initializes a new instance of the <see cref="ChartEx"/> class.
        /// </summary>
        public ChartEx() : base(ElementType.Chart) { }

        #endregion
    }
}

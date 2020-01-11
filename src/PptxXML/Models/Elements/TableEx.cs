using DocumentFormat.OpenXml;
using PptxXML.Enums;

namespace PptxXML.Models.Elements
{
    /// <summary>
    /// Represents a table element on a slide.
    /// </summary>
    public class TableEx: Element
    {
        #region Constructors

        /// <summary>
        /// Initialise an instance of <see cref="TableEx"/> class.
        /// </summary>
        public TableEx(OpenXmlCompositeElement ce) : base(ElementType.Table, ce) { }

        #endregion Constructors
    }
}
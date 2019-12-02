using DocumentFormat.OpenXml;
using PptxXML.Enums;

namespace PptxXML.Entities.Elements
{
    /// <summary>
    /// Represent a table on a slide.
    /// </summary>
    public class Table: Element
    {
        #region Constructors

        /// <summary>
        /// Initialise an instance of <see cref="Table"/> class.
        /// </summary>
        /// <param name="xmlCompositeElement"></param>
        public Table(OpenXmlCompositeElement xmlCompositeElement) :
            base(xmlCompositeElement)
        {
            Type = ElementType.Table;
        }

        #endregion
    }
}
using DocumentFormat.OpenXml;

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
        /// <param name="xmlCompositeElement"></param>
        public TableEx(OpenXmlCompositeElement xmlCompositeElement) : base(xmlCompositeElement)
        {

        }

        #endregion Constructors
    }
}
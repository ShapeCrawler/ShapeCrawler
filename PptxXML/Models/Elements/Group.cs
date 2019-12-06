using DocumentFormat.OpenXml;
using PptxXML.Enums;

namespace PptxXML.Models.Elements
{
    /// <summary>
    /// Represent a group shape.
    /// </summary>
    public class Group: Element
    {
        #region Constructors

        /// <summary>
        /// Initialise an instance of <see cref="Group"/> class.
        /// </summary>
        /// <param name="xmlCompositeElement"></param>
        public Group(OpenXmlCompositeElement xmlCompositeElement) : base(xmlCompositeElement)
        {
            Type = ElementType.Group;
        }

        #endregion
    }
}

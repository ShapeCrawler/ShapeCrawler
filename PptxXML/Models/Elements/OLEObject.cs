using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using PptxXML.Enums;

namespace PptxXML.Entities.Elements
{
    /// <summary>
    /// Represent OLE (Object Linking and Embedding) object.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class OLEObject: Element
    {
        #region Constructors

        /// <summary>
        /// Initialise an instance of <see cref="OLEObject"/> class.
        /// </summary>
        /// <param name="xmlCompositeElement"></param>
        public OLEObject(OpenXmlCompositeElement xmlCompositeElement) :
            base(xmlCompositeElement)
        {
            Type = ElementType.OLEObject;
        }

        #endregion
    }
}

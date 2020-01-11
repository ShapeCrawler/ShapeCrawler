using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using PptxXML.Enums;

namespace PptxXML.Models.Elements
{
    /// <summary>
    /// Represents an OLE object on a slide.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class OLEObject: Element
    {
        #region Constructors

        /// <summary>
        /// Initialise a new instance of the <see cref="OLEObject"/> class.
        /// </summary>
        public OLEObject(OpenXmlCompositeElement compositeElement) : base(ElementType.OLEObject, compositeElement)
        {

        }

        #endregion
    }
}

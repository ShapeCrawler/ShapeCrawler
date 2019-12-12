using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using PptxXML.Enums;

namespace PptxXML.Models.Elements
{
    /// <summary>
    /// Represents a OLE (Object Linking and Embedding) object.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class OLEObject: Element
    {
        #region Constructors

        /// <summary>
        /// Initialise a new instance of the <see cref="OLEObject"/> class.
        /// </summary>
        public OLEObject() : base(ElementType.OLEObject) { }

        #endregion
    }
}

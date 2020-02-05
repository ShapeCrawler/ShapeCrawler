using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using SlideXML.Validation;

namespace SlideXML.Models.SlideComponents
{
    /// <summary>
    /// Represents an OLE object on a slide.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class OleObjectSL
    {
        private readonly OpenXmlCompositeElement _compositeElement;

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="OleObjectSL"/> class.
        /// </summary>
        public OleObjectSL(OpenXmlCompositeElement ce)
        {
            Check.NotNull(ce, nameof(ce));
            _compositeElement = ce;
        }

        #endregion
    }
}

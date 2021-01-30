using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shared;

namespace ShapeCrawler.Models.SlideComponents
{
    /// <summary>
    /// Represents an OLE object on a slide.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class OLEObject
    {
        private readonly OpenXmlCompositeElement _compositeElement;

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="OLEObject"/> class.
        /// </summary>
        public OLEObject(OpenXmlCompositeElement ce)
        {
            Check.NotNull(ce, nameof(ce));
            _compositeElement = ce;
        }

        #endregion
    }
}

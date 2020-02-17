using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using SlideDotNet.Validation;

namespace SlideDotNet.Models.SlideComponents
{
    /// <summary>
    /// Represents an OLE object on a slide.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class OleObject
    {
        private readonly OpenXmlCompositeElement _compositeElement;

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="OleObject"/> class.
        /// </summary>
        public OleObject(OpenXmlCompositeElement ce)
        {
            Check.NotNull(ce, nameof(ce));
            _compositeElement = ce;
        }

        #endregion
    }
}

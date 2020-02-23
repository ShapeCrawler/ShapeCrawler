using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Validation;

namespace SlideDotNet.Models.SlideComponents
{
    /// <summary>
    /// Represents a picture content.
    /// </summary>
    public class Picture
    {
        #region Fields

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets image.
        /// </summary>
        public ImageEx ImageEx { get; }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Picture"/> class.
        /// </summary>
        public Picture(SlidePart xmlSldPart, string blipRelateId)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));
            ImageEx = new ImageEx(xmlSldPart, blipRelateId);
        }

        #endregion Constructors
    }
}
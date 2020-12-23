using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Services.Drawing;
using ShapeCrawler.Shared;

namespace ShapeCrawler.Models.SlideComponents
{
    /// <summary>
    /// Represents a picture content.
    /// </summary>
    public class PictureEx
    {
        #region Properties

        /// <summary>
        /// Gets image.
        /// </summary>
        public ImageEx ImageEx { get; }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="PictureEx"/> class.
        /// </summary>
        public PictureEx(SlidePart xmlSldPart, string blipRelateId)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));
            ImageEx = new ImageEx(xmlSldPart, blipRelateId);
        }

        #endregion Constructors
    }
}
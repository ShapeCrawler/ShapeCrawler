using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Factories.Drawing;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Models.SlideComponents
{
    /// <summary>
    /// Represents a picture content.
    /// </summary>
    public class PictureSc
    {
        private readonly P.Picture _pPicture;

        #region Properties

        /// <summary>
        /// Gets image.
        /// </summary>
        public ImageSc Image { get; }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="PictureSc"/> class.
        /// </summary>
        public PictureSc(SlidePart slidePart, string blipRelateId)
        {
            Check.NotNull(slidePart, nameof(slidePart));
            Image = new ImageSc(slidePart, blipRelateId);
        }

        #endregion Constructors
    }
}
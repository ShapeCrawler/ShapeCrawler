using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Enums;
using ShapeCrawler.Factories.Drawing;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMaster;
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

        public PictureSc(P.Picture pPicture)
        {
            _pPicture = pPicture;
        }

        #endregion Constructors
    }

    public class PictureScNew : BaseShape
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
        public PictureScNew(SlidePart xmlSldPart, string blipRelateId)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));
            Image = new ImageSc(xmlSldPart, blipRelateId);
        }

        public PictureScNew(SlideMasterSc slideMaster, P.Picture pPicture) : base(slideMaster, pPicture)
        {
            _pPicture = pPicture;
        }

        #endregion Constructors

        public override long X { get; }
        public override long Y { get; }
        public override long Width { get; }
        public override long Height { get; }
        public override GeometryType GeometryType { get; }
    }
}
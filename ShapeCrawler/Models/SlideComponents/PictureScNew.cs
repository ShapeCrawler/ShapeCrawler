using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Factories.Drawing;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler.Models.SlideComponents
{
    public class PictureScNew : BaseShape
    {
        private readonly DocumentFormat.OpenXml.Presentation.Picture _pPicture;

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

        public PictureScNew(SlideMasterSc slideMaster, DocumentFormat.OpenXml.Presentation.Picture pPicture) : base(slideMaster, pPicture)
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
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Enums;
using ShapeCrawler.Factories.Drawing;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Models.SlideComponents
{
    /// <summary>
    /// Represents a picture content.
    /// </summary>
    public class PictureSc : BaseShape
    {
        private readonly P.Picture _pPicture;

        #region Properties

        /// <summary>
        /// Gets image.
        /// </summary>
        public ImageEx ImageEx { get; }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="PictureSc"/> class.
        /// </summary>
        public PictureSc(SlidePart xmlSldPart, string blipRelateId)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));
            ImageEx = new ImageEx(xmlSldPart, blipRelateId);
        }

        public PictureSc(SlideMaster.SlideMasterSc slideMaster, P.Picture pPicture) : base(pPicture)
        {
            _pPicture = pPicture;
        }

        #endregion Constructors

        public override long Width => throw new System.NotImplementedException();

        public override long Height => throw new System.NotImplementedException();

        public override long X => throw new System.NotImplementedException();

        public override long Y => throw new System.NotImplementedException();

        public override GeometryType GeometryType => throw new System.NotImplementedException();
    }
}
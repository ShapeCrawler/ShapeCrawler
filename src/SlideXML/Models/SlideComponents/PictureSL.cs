using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using LogicNull.Utilities;
using SlideXML.Exceptions;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Models.SlideComponents
{
    /// <summary>
    /// Represents a picture element.
    /// </summary>
    public class PictureSL
    {
        #region Fields

        private readonly SlidePart _sldPart;
        private ImageEx _imageEx;
        private readonly OpenXmlCompositeElement _compositeElement;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets image.
        /// </summary>
        public ImageEx ImageEx
        {
            get
            {
                if (_imageEx == null)
                {
                    var pPicture = (P.Picture)_compositeElement;
                    var pBlipFill = pPicture.GetFirstChild<P.BlipFill>();
                    var blipRelateId = pBlipFill?.Blip?.Embed?.Value;
                    if (blipRelateId != null)
                    {
                        _imageEx = new ImageEx(_sldPart, blipRelateId);
                    }
                    else
                    {
                        throw new SlideXMLException("Element does contain an image.");
                    }
                }

                return _imageEx;
            }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of <see cref="PictureSL"/> class.
        /// </summary>
        public PictureSL(SlidePart sldPart, OpenXmlCompositeElement compositeElement)
        {
            Check.NotNull(sldPart, nameof(sldPart));
            _sldPart = sldPart;
            _compositeElement = compositeElement;
        }

        #endregion Constructors
    }
}
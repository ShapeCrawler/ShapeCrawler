using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using LogicNull.Utilities;
using SlideXML.Enums;
using SlideXML.Exceptions;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Models.Elements
{
    /// <summary>
    /// Represents a picture element.
    /// </summary>
    public class PictureEx: Element
    {
        #region Fields

        private readonly SlidePart _sldPart;
        private ImageEx _imageEx;

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
                    var pPicture = (P.Picture)CompositeElement;
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
        /// Initializes a new instance of <see cref="PictureEx"/> class.
        /// </summary>
        public PictureEx(SlidePart sldPart, OpenXmlCompositeElement compositeElement) : base(ElementType.Picture, compositeElement)
        {
            Check.NotNull(sldPart, nameof(sldPart));
            _sldPart = sldPart;
        }

        #endregion Constructors
    }
}
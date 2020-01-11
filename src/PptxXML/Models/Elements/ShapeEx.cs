using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ObjectEx.Utilities;
using PptxXML.Enums;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Models.Elements
{
    /// <summary>
    /// Represents a shape element on a slide.
    /// </summary>
    public class ShapeEx : Element
    {
        #region Fields

        private readonly SlidePart _sldPart;

        private bool _xmlTxtBodyParsed; // used to avoid second time parsed text body
        private TextBodyEx _textBody;
        private ImageEx _backgroundImage;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets text body.
        /// </summary>
        /// <remarks>Lazy load.</remarks>
        public TextBodyEx TextBody
        {
            get
            {
                if (!_xmlTxtBodyParsed)
                {
                    _textBody = TryParseTxtBody();
                }

                return _textBody;
            }
        }

        /// <summary>
        /// Gets background image.
        /// </summary>
        /// <returns><see cref="ImageEx"/> instance or null if shape has not background image.</returns>
        public ImageEx BackgroundImage
        {
            get
            {
                if (_backgroundImage == null)
                {
                    var pShape = (P.Shape)CompositeElement;
                    var aBlipFill = pShape.ShapeProperties.GetFirstChild<A.BlipFill>();
                    var blipRelateId = aBlipFill?.Blip?.Embed?.Value; // try to get blip relationship ID
                    if (blipRelateId != null)
                    {
                        _backgroundImage = new ImageEx(_sldPart, blipRelateId);
                    }
                }

                return _backgroundImage;
            }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ShapeEx"/> class.
        /// </summary>
        [SuppressMessage("ReSharper", "SuggestBaseTypeForParameter")]
        public ShapeEx(OpenXmlCompositeElement compositeElement, SlidePart sldPart) : base(ElementType.Shape, compositeElement)
        {
            Check.NotNull(sldPart, nameof(sldPart));
            _sldPart = sldPart;
        }

        #endregion Constructors

        #region Private Methods

        private TextBodyEx TryParseTxtBody()
        {
            // TextBodyEx
            TextBodyEx result = null;
            var xmlTxtBody = ((P.Shape)CompositeElement).TextBody;
            if (xmlTxtBody != null)
            {
                var aTexts = xmlTxtBody.Descendants<A.Text>();
                if (aTexts.Any(t => t.Parent is A.Run)
                    && aTexts.Sum(t => t.Text.Length) > 0) // at least one of <a:t> element contain text
                {
                    result = new TextBodyEx(xmlTxtBody);
                }
            }

            _xmlTxtBodyParsed = true;
            return result; // if shape does have text null is returned
        }

        #endregion
    }
}
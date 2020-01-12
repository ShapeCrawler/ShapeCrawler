using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ObjectEx.Utilities;
using PptxXML.Enums;
using PptxXML.Models.Elements.Builders;
using PptxXML.Services;
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

        #region Dependencies

        private readonly IBackgroundImageFactory _bgImgFactory;

        #endregion Dependencies

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
                return _backgroundImage ??= _bgImgFactory.CreateBackgroundShape(_sldPart, (P.Shape) CompositeElement);
            }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ShapeEx"/> class.
        /// </summary>
        [SuppressMessage("ReSharper", "SuggestBaseTypeForParameter")]
        private ShapeEx(OpenXmlCompositeElement compositeElement, SlidePart sldPart, IBackgroundImageFactory bgImgFactory) : base(ElementType.Shape, compositeElement)
        {
            Check.NotNull(sldPart, nameof(sldPart));
            Check.NotNull(bgImgFactory, nameof(bgImgFactory));
            _sldPart = sldPart;
            _bgImgFactory = bgImgFactory;
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

        #region Builder

        /// <summary>
        /// Represents <see cref="ShapeEx"/> instance builder.
        /// </summary>
        public class Builder : IShapeExBuilder
        {
            #region Dependencies

            private readonly IBackgroundImageFactory _bgImgFactor;

            #endregion Dependencies

            #region Constructors

            public Builder(IBackgroundImageFactory bgImgFactor)
            {
                Check.NotNull(bgImgFactor, nameof(bgImgFactor));
                _bgImgFactor = bgImgFactor;
            }

            #endregion Constructors

            #region Public Methods

            /// <summary>
            /// Builds shape.
            /// </summary>
            /// <returns></returns>
            public ShapeEx Build(OpenXmlCompositeElement compositeElement, SlidePart sldPart)
            {
                Check.NotNull(compositeElement, nameof(compositeElement));
                Check.NotNull(sldPart, nameof(sldPart));

                return new ShapeEx(compositeElement, sldPart, _bgImgFactor);
            }

            #endregion Public Methods
        }

        #endregion Builder
    }
}
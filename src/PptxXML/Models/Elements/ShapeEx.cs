using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ObjectEx.Utilities;
using PptxXML.Enums;
using PptxXML.Models.Settings;
using PptxXML.Models.TextBody;
using PptxXML.Services;
using PptxXML.Services.Builders;
using PptxXML.Services.Placeholder;
using PptxXML.Services.Placeholders;
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
        private ShapeSettings _spSettings;

        #endregion Fields

        #region Dependencies

        private readonly IBackgroundImageFactory _bgImgFactory;
        private readonly ITextBodyExBuilder _txtBodyBuilder;

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
        private ShapeEx(IBackgroundImageFactory bgImgFactory, 
                        ITextBodyExBuilder txtBodyBuilder,
                        OpenXmlCompositeElement compositeElement,
                        SlidePart sldPart,
                        ShapeSettings spSettings) : base(ElementType.Shape, compositeElement)
        {
            Check.NotNull(bgImgFactory, nameof(bgImgFactory));
            Check.NotNull(txtBodyBuilder, nameof(txtBodyBuilder));
            Check.NotNull(sldPart, nameof(sldPart));
            Check.NotNull(spSettings, nameof(spSettings));
            _bgImgFactory = bgImgFactory;
            _txtBodyBuilder = txtBodyBuilder;
            _sldPart = sldPart;
            _spSettings = spSettings;
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
                    result = _txtBodyBuilder.Build(xmlTxtBody, _spSettings);
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
            private readonly ITextBodyExBuilder _txtBodyBuilder;

            #endregion Dependencies

            #region Constructors

            public Builder(IBackgroundImageFactory bgImgFactor, ITextBodyExBuilder txtBodyBuilder)
            {
                Check.NotNull(bgImgFactor, nameof(bgImgFactor));
                Check.NotNull(txtBodyBuilder, nameof(txtBodyBuilder));
                _bgImgFactor = bgImgFactor;
                _txtBodyBuilder = txtBodyBuilder;
            }

            #endregion Constructors

            #region Public Methods

            /// <summary>
            /// Builds shape.
            /// </summary>
            /// <returns></returns>
            public ShapeEx Build(OpenXmlCompositeElement compositeElement, SlidePart sldPart, ShapeSettings spSettings)
            {
                Check.NotNull(compositeElement, nameof(compositeElement));
                Check.NotNull(sldPart, nameof(sldPart));

                return new ShapeEx(_bgImgFactor, _txtBodyBuilder, compositeElement, sldPart, spSettings);
            }

            #endregion Public Methods
        }

        #endregion Builder
    }
}
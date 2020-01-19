using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using LogicNull.Utilities;
using SlideXML.Enums;
using SlideXML.Models.Settings;
using SlideXML.Models.TextBody;
using SlideXML.Services;
using SlideXML.Services.Builders;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideXML.Models.Elements
{
    /// <summary>
    /// Represents a shape element on a slide.
    /// </summary>
    public class ShapeEx : Element
    {
        #region Fields

        private readonly SlidePart _sldPart;
        private readonly ElementSettings _spSettings;

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
                if (_textBody == null)
                {
                    TryParseTxtBody();
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
                        OpenXmlCompositeElement compositeElement,
                        SlidePart sldPart,
                        ElementSettings spSettings) : base(ElementType.Shape, compositeElement)
        {
            Check.NotNull(bgImgFactory, nameof(bgImgFactory));
            Check.NotNull(sldPart, nameof(sldPart));
            Check.NotNull(spSettings, nameof(spSettings));
            _bgImgFactory = bgImgFactory;
            _sldPart = sldPart;
            _spSettings = spSettings;
        }

        #endregion Constructors

        #region Private Methods

        private void TryParseTxtBody()
        {
            var pTxtBody = ((P.Shape)CompositeElement).TextBody;

            if (pTxtBody == null)
            {
                return;
            }

            var aTexts = pTxtBody.Descendants<A.Text>();
            if (aTexts.Any(t => t.Parent is A.Run) && aTexts.Sum(t => t.Text.Length) > 0) // at least one of <a:t> element contain text
            {
                _textBody = new TextBodyEx(_spSettings, pTxtBody);
            }
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
            public ShapeEx Build(OpenXmlCompositeElement compositeElement, SlidePart sldPart, ElementSettings spSettings)
            {
                Check.NotNull(compositeElement, nameof(compositeElement));
                Check.NotNull(sldPart, nameof(sldPart));

                return new ShapeEx(_bgImgFactor, compositeElement, sldPart, spSettings);
            }

            #endregion Public Methods
        }

        #endregion Builder
    }
}
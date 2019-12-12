using System.Diagnostics.CodeAnalysis;
using System.Linq;
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

        private bool _xmlTxtBodyParsed; // used to avoid second time parsed text body
        private TextBodyEx _textBody;

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

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ShapeEx"/> class.
        /// </summary>
        [SuppressMessage("ReSharper", "SuggestBaseTypeForParameter")]
        public ShapeEx() : base(ElementType.Shape) { }

        #endregion Constructors

        #region Private Methods

        private TextBodyEx TryParseTxtBody()
        {
            // TextBodyEx
            TextBodyEx result = null;
            var xmlTxtBody = ((P.Shape)XmlCompositeElement).TextBody;
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
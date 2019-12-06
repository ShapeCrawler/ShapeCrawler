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

        private readonly P.Shape _xmlShape;
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
        /// <param name="xmlShape"></param>
        [SuppressMessage("ReSharper", "SuggestBaseTypeForParameter")]
        public ShapeEx(P.Shape xmlShape) : base(xmlShape)
        {
            _xmlShape = xmlShape;
            Init();
        }

        #endregion Constructors

        #region Private Methods

        private void Init()
        {
            // Type
            Type = ElementType.Shape;

            // X, Y, W, H
            var t2D = _xmlShape.ShapeProperties.Transform2D;
            X = t2D.Offset.X.Value;
            Y = t2D.Offset.Y.Value;
            Width = t2D.Extents.Cx.Value;
            Height = t2D.Extents.Cy.Value;
        }

        private TextBodyEx TryParseTxtBody()
        {
            // TextBodyEx
            TextBodyEx result = null;
            var xmlTxtBody = _xmlShape.TextBody;
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
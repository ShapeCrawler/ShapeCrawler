using System.Diagnostics.CodeAnalysis;
using System.Linq;
using PptxXML.Entities.Elements;
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

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets text body.
        /// </summary>
        public TextBodyEx TextBody { get; private set; }

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

            // TextBodyEx
            var xmlTxtBody = _xmlShape.TextBody;
            if (xmlTxtBody != null)
            {
                var aTexts = xmlTxtBody.Descendants<A.Text>();
                if (aTexts.Any(t => t.Parent is A.Run) 
                    && aTexts.Sum(t => t.Text.Length) > 0) // at least one of <a:t> element contain text
                {
                    TextBody = new TextBodyEx(xmlTxtBody);
                }
            }
        }

        #endregion
    }
}
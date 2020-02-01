using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using LogicNull.Utilities;
using SlideXML.Models.Settings;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideXML.Models.TextBody
{
    /// <summary>
    /// Represents a text body of the shape.
    /// </summary>
    public class TextFrame
    {
        #region Fields

        private readonly ElementSettings _spSettings;
        private string _text;

        #endregion

        #region Properties

        /// <summary>
        /// Gets paragraphs.
        /// </summary>
        public IList<ParagraphSL> Paragraphs { get; private set; }

        public string Text
        {
            get
            {
                if (_text == null)
                {
                    InitText();
                }

                return _text;
            }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes an instance of the <see cref="TextFrame"/>.
        /// </summary>
        public TextFrame(ElementSettings spSettings, OpenXmlCompositeElement compositeElement)
        {
            Check.NotNull(spSettings, nameof(spSettings));
            Check.NotNull(compositeElement, nameof(compositeElement));
            _spSettings = spSettings;
            ParseParagraphs(compositeElement);
        }

        #endregion Constructors

        #region Private Methods

        private void ParseParagraphs(OpenXmlCompositeElement compositeElement)
        {
            // Parses non-empty paragraphs
            var paragraphs = compositeElement.Elements<A.Paragraph>().Where(e => e.Descendants<A.Text>().Any());

            // Sets paragraphs
            Paragraphs = new List<ParagraphSL>(paragraphs.Count());
            foreach (var p in paragraphs)
            {
                Paragraphs.Add(new ParagraphSL(_spSettings, p));
            }
        }

        private void InitText()
        {
            var sb = new StringBuilder();
            sb.Append(Paragraphs[0].Text);
            
            // If the number of paragraphs more than one.
            var numPr = Paragraphs.Count;
            var index = 1;
            while (index < numPr)
            {
                sb.AppendLine();
                sb.Append(Paragraphs[index].Text);

                index++;
            }

            _text = sb.ToString();
        }

        #endregion Private Methods
    }
}

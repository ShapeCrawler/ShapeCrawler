using System.Collections.Generic;
using System.Linq;
using System.Text;
using LogicNull.Utilities;
using SlideXML.Models.Settings;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideXML.Models.TextBody
{
    /// <summary>
    /// Represents a text body of the shape.
    /// </summary>
    public class TextBodyEx
    {
        #region Fields

        private readonly ElementSettings _spSettings;
        private string _text;

        #endregion

        #region Properties

        /// <summary>
        /// Gets paragraphs.
        /// </summary>
        public IList<ParagraphEx> Paragraphs { get; private set; }

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
        /// Initializes an instance of the <see cref="TextBodyEx"/> class with <see cref="P.TextBody"/>.
        /// </summary>
        /// <param name="spSettings"></param>
        /// <param name="pTxtBody"><see cref="P.TextBody"/> instance which contains a text.</param>
        public TextBodyEx(ElementSettings spSettings, P.TextBody pTxtBody)
        {
            Check.NotNull(spSettings, nameof(spSettings));
            Check.NotNull(pTxtBody, nameof(pTxtBody));
            _spSettings = spSettings;
            ParseParagraphs(pTxtBody);
        }

        /// <summary>
        /// Initializes an instance of the <see cref="TextBodyEx"/> class with <see cref="A.TextBody"/>.
        /// </summary>
        /// <param name="spSettings"></param>
        /// <param name="aTxtBody"><see cref="A.TextBody"/> instance which contains a text.</param>
        public TextBodyEx(ElementSettings spSettings, A.TextBody aTxtBody)
        {
            Check.NotNull(spSettings, nameof(spSettings));
            Check.NotNull(spSettings, nameof(aTxtBody));
            _spSettings = spSettings;
            ParseParagraphs(aTxtBody);
        }

        #endregion Constructors

        #region Private Methods

        private void ParseParagraphs(P.TextBody pTxtBody)
        {
            var aParagraphs = pTxtBody.Elements<A.Paragraph>();
            SetParagraphs(aParagraphs);
        }

        private void ParseParagraphs(A.TextBody aTxtBody)
        {
            var aParagraphs = aTxtBody.Elements<A.Paragraph>();
            SetParagraphs(aParagraphs);
        }

        private void SetParagraphs(IEnumerable<A.Paragraph> paragraphs)
        {
            Paragraphs = new List<ParagraphEx>(paragraphs.Count());
            foreach (var p in paragraphs)
            {
                Paragraphs.Add(new ParagraphEx(_spSettings, p));
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

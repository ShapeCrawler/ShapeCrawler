using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using SlideDotNet.Models.Settings;
using SlideDotNet.Validation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Models.TextBody
{
    /// <summary>
    /// <inheritdoc cref="ITextFrame"/>
    /// </summary>
    public sealed class TextFrame : ITextFrame
    {
        #region Fields

        private readonly ElementSettings _elSettings;
        private string _text;

        #endregion

        #region Properties

        public IList<Paragraph> Paragraphs { get; private set; }

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
        public TextFrame(ElementSettings elSettings, OpenXmlCompositeElement compositeElement)
        {
            _elSettings = elSettings ?? throw new ArgumentNullException(nameof(elSettings));
            Check.NotNull(compositeElement, nameof(compositeElement));
            ParseParagraphs(compositeElement);
        }

        #endregion Constructors

        #region Private Methods

        private void ParseParagraphs(OpenXmlCompositeElement compositeElement)
        {
            // Parses non-empty paragraphs
            var paragraphs = compositeElement.Elements<A.Paragraph>().Where(e => e.Descendants<A.Text>().Any());

            // Sets paragraphs
            Paragraphs = new List<Paragraph>(paragraphs.Count());
            foreach (var p in paragraphs)
            {
                Paragraphs.Add(new Paragraph(_elSettings, p));
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

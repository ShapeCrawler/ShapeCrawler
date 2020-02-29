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

        private readonly IShapeContext _spContext;
        private readonly Lazy<string> _text;

        #endregion Fields

        #region Properties

        /// <summary>
        /// <inheritdoc cref="ITextFrame.Paragraphs"/>
        /// </summary>
        public IList<Paragraph> Paragraphs { get; private set; }

        /// <summary>
        /// <inheritdoc cref="ITextFrame.Text"/>
        /// </summary>
        public string Text => _text.Value;

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes an instance of the <see cref="TextFrame"/>.
        /// </summary>
        public TextFrame(IShapeContext spContext, OpenXmlCompositeElement compositeElement)
        {
            _spContext = spContext ?? throw new ArgumentNullException(nameof(spContext));
            Check.NotNull(compositeElement, nameof(compositeElement));
            ParseParagraphs(compositeElement);
            _text = new Lazy<string>(GetText);
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
                Paragraphs.Add(new Paragraph(_spContext, p));
            }
        }

        private string GetText()
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

            return sb.ToString();
        }

        #endregion Private Methods
    }
}

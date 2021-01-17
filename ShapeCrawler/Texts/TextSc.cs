using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Settings;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts
{
    // TODO: Override ToString()
    /// <summary>
    /// <inheritdoc cref="ITextFrame"/>
    /// </summary>
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    public sealed class TextSc : ITextFrame
    {
        #region Fields

        private readonly ShapeContext _spContext;
        private readonly Lazy<string> _text;

        #endregion Fields

        internal ShapeSc ShapeEx { get; }

        #region Public Properties

        /// <summary>
        /// <inheritdoc cref="ITextFrame.Paragraphs"/>
        /// </summary>
        public IList<ParagraphSc> Paragraphs { get; private set; } // TODO: Consider to use IReadOnlyList instead IList or create own collection

        /// <summary>
        /// <inheritdoc cref="ITextFrame.Text"/>
        /// </summary>
        public string Text => _text.Value;

        #endregion Public Properties

        #region Constructors

        /// <summary>
        /// Initializes an instance of the <see cref="TextSc"/>.
        /// </summary>
        public TextSc(OpenXmlCompositeElement compositeElement, ShapeSc shapeEx)
            :this(shapeEx.Context, compositeElement)
        {
            ShapeEx = shapeEx;
        }

        public TextSc(ShapeContext spContext, OpenXmlCompositeElement compositeElement)
        {
            _spContext = spContext;
            ParseParagraphs(compositeElement); // TODO: Make paragraphs parsing lazy
            _text = new Lazy<string>(GetText);
        }

        #endregion Constructors

        #region Private Methods

        private void ParseParagraphs(OpenXmlCompositeElement compositeElement)
        {
            // Parses non-empty paragraphs
            var aParagraphs = compositeElement.Elements<A.Paragraph>().Where(e => e.Descendants<A.Text>().Any());

            // Sets paragraphs
            Paragraphs = new List<ParagraphSc>(aParagraphs.Count());
            foreach (A.Paragraph aParagraph in aParagraphs)
            {
                Paragraphs.Add(new ParagraphSc(_spContext, aParagraph, this));
            }
        }

        private string GetText()
        {
            var sb = new StringBuilder();
            sb.Append(Paragraphs[0].Text);
            
            // If the number of paragraphs more than one
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

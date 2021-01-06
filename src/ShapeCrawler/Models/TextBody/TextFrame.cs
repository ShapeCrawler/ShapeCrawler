using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Settings;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Models.TextBody
{
    /// <summary>
    /// <inheritdoc cref="ITextFrame"/>
    /// </summary>
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    public sealed class TextFrame : ITextFrame
    {
        #region Fields

        private readonly IShapeContext _spContext;
        private readonly Lazy<string> _text;

        #endregion Fields

        internal Shape Shape { get; }


        #region Public Properties

        /// <summary>
        /// <inheritdoc cref="ITextFrame.Paragraphs"/>
        /// </summary>
        public IList<Paragraph> Paragraphs { get; private set; } // TODO: Consider to use IReadOnlyList instead IList

        /// <summary>
        /// <inheritdoc cref="ITextFrame.Text"/>
        /// </summary>
        public string Text => _text.Value;

        #endregion Public Properties

        #region Constructors

        /// <summary>
        /// Initializes an instance of the <see cref="TextFrame"/>.
        /// </summary>
        public TextFrame(IShapeContext spContext, OpenXmlCompositeElement compositeElement, Shape shape)
        :this(spContext, compositeElement)
        {
            Shape = shape;
        }

        public TextFrame(IShapeContext spContext, OpenXmlCompositeElement compositeElement)
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
            Paragraphs = new List<Paragraph>(aParagraphs.Count());
            foreach (A.Paragraph aParagraph in aParagraphs)
            {
                Paragraphs.Add(new Paragraph(_spContext, aParagraph, this));
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

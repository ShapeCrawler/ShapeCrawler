using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    internal sealed class SCTextBox : ITextBox
    {
        #region Fields

        private readonly Lazy<string> _text;
        private readonly OpenXmlCompositeElement _compositeElement;
        private ParagraphCollection _paragraphs;
        internal Shape AutoShape { get; }

        #endregion Fields

        #region Public Properties

        /// <summary>
        ///     Gets text paragraph collection.
        /// </summary>
        public ParagraphCollection Paragraphs => _paragraphs;

        /// <summary>
        ///     Gets or sets text box string content. Returns null if the text box is empty.
        /// </summary>
        public string Text
        {
            get => _text.Value;
            set => SetText(value);
        }

        #endregion Public Properties

        #region Constructors

        internal SCTextBox(Shape autoShape, P.TextBody pTextBody)
        {
            AutoShape = autoShape;
            _compositeElement = pTextBody;
            _text = new Lazy<string>(GetText);
            _paragraphs = new ParagraphCollection(_compositeElement, this);
        }

        // TODO: Resolve conflict getting text box for autoShape and table
        internal SCTextBox(A.TextBody aTextBody)
        {
            _compositeElement = aTextBody;
            _text = new Lazy<string>(GetText);
            _paragraphs = new ParagraphCollection(_compositeElement, this);
        }

        #endregion Constructors

        #region Private Methods

        private void SetText(string value)
        {
            SCParagraph paragraph = Paragraphs.First(p => p.Portions != null);
            foreach (SCParagraph removingPara in Paragraphs.Where(p => p != paragraph))
            {
                removingPara.AParagraph.Remove();
            }
            _paragraphs = new ParagraphCollection(_compositeElement, this);            

            Paragraphs.Single().Text = value;
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
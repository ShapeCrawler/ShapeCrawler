using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using ShapeCrawler.Collections;
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
        private readonly OpenXmlCompositeElement _textBodyCompositeElement;
        internal Shape AutoShape { get; }

        #endregion Fields

        #region Public Properties

        /// <summary>
        ///     Gets collection of text paragraphs.
        /// </summary>
        public IParagraphCollection Paragraphs { get; private set; }

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

        internal SCTextBox(P.TextBody autoShapePTextBody, Shape autoShape)
        {
            _textBodyCompositeElement = autoShapePTextBody;
            _text = new Lazy<string>(GetText);
            Paragraphs = new ParagraphCollection(_textBodyCompositeElement, this);

            AutoShape = autoShape;
        }

        internal SCTextBox(A.TextBody tblCellATextBody)
        {
            _textBodyCompositeElement = tblCellATextBody;
            _text = new Lazy<string>(GetText);
            Paragraphs = new ParagraphCollection(_textBodyCompositeElement, this);
        }

        #endregion Constructors

        #region Private Methods

        private void SetText(string value)
        {
            IParagraph paragraph = Paragraphs.First(p => p.Portions != null);
            foreach (SCParagraph removingPara in Paragraphs.Where(p => p != paragraph))
            {
                removingPara.AParagraph.Remove();
            }

            Paragraphs = new ParagraphCollection(_textBodyCompositeElement, this);

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
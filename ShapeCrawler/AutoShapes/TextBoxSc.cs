using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using ShapeCrawler.Settings;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes
{
    // TODO: Override ToString()

    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    public sealed class TextBoxSc : ITextBox
    {
        #region Fields

        private readonly Lazy<string> _text;
        private readonly OpenXmlCompositeElement _compositeElement;

        internal Shape AutoShape { get; }

        #endregion Fields

        #region Public Properties

        /// <summary>
        ///     Gets text paragraph collection.
        /// </summary>
        public ParagraphCollection Paragraphs => new(_compositeElement, this);

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

        internal TextBoxSc(Shape autoShape, P.TextBody pTextBody)
        {
            AutoShape = autoShape;
            _compositeElement = pTextBody;
            _text = new Lazy<string>(GetText);
        }

        // TODO: Resolve conflict getting text box for autoShape and table
        internal TextBoxSc(A.TextBody aTextBody)
        {
            _compositeElement = aTextBody;
            _text = new Lazy<string>(GetText);
        }

        #endregion Constructors

        #region Private Methods

        private void SetText(string value)
        {
            if (Paragraphs.Count > 1)
            {
                // Remove all except first paragraph
                Paragraphs.RemoveRange(1, Paragraphs.Count - 1);
            }

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
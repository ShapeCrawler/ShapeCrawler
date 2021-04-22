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
    /// <summary>
    ///     <inheritdoc cref="ITextBox"/>
    /// </summary>
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    internal sealed class SCTextBox : ITextBox
    {
        #region Fields

        private readonly Lazy<string> text;
        private readonly OpenXmlCompositeElement textBodyCompositeElement;

        /// <summary>
        ///     Initializes a new instance of the <see cref="SCTextBox"/> class for Auto Shape.
        /// </summary>
        internal SCTextBox(P.TextBody autoShapePTextBody, Shape autoShape)
        {
            this.textBodyCompositeElement = autoShapePTextBody;
            this.text = new Lazy<string>(this.GetText);
            this.Paragraphs = new ParagraphCollection(this.textBodyCompositeElement, this);
            this.ParentAutoShape = autoShape;
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="SCTextBox"/> class for Table Cell.
        /// </summary>
        internal SCTextBox(A.TextBody tblCellATextBody)
        {
            this.textBodyCompositeElement = tblCellATextBody;
            this.text = new Lazy<string>(this.GetText);
            this.Paragraphs = new ParagraphCollection(this.textBodyCompositeElement, this);
        }

        /// <inheritdoc/>
        public IParagraphCollection Paragraphs { get; private set; }

        /// <inheritdoc/>
        public string Text
        {
            get => this.text.Value;
            set => this.SetText(value);
        }

        internal Shape ParentAutoShape { get; }

        private void SetText(string value)
        {
            bool changed = false;
            SCParagraph paragraph = (SCParagraph)this.Paragraphs.First(p => p.Portions.Any());
            foreach (SCParagraph removingPara in this.Paragraphs.Where(p => p != paragraph))
            {
                removingPara.AParagraph.Remove();
                changed = true;
            }

            if (changed)
            {
                this.Paragraphs = new ParagraphCollection(this.textBodyCompositeElement, this);
            }

            this.Paragraphs.Single().Text = value;
        }

        private string GetText()
        {
            var sb = new StringBuilder();
            sb.Append(this.Paragraphs[0].Text);

            // If the number of paragraphs more than one
            var numPr = this.Paragraphs.Count;
            var index = 1;
            while (index < numPr)
            {
                sb.AppendLine();
                sb.Append(this.Paragraphs[index].Text);

                index++;
            }

            return sb.ToString();
        }

        #endregion Private Methods

        public void ThrowIfRemoved()
        {
            this.ParentAutoShape.ThrowIfRemoved();
        }
    }
}
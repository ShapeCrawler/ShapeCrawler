using System;
using System.Collections.Generic;
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
    internal class SCTextBox : ITextBox
    {
        private readonly Lazy<string> text;
        private readonly OpenXmlCompositeElement textBodyCompositeElement; // instance of A.TextBody or P.TextBody class

        internal SCTextBox(OpenXmlCompositeElement textBodyCompositeElement, ITextBoxContainer parentTextBoxContainer)
        {
            this.text = new Lazy<string>(this.GetText);
            this.textBodyCompositeElement = textBodyCompositeElement;
            this.Paragraphs = new ParagraphCollection(this.textBodyCompositeElement, this);
            this.ParentTextBoxContainer = parentTextBoxContainer;
        }

        #region Public Properties

        public IParagraphCollection Paragraphs { get; private set; }

        public string Text
        {
            get => this.text.Value;
            set => this.SetText(value);
        }

        #endregion Public Properties

        internal ITextBoxContainer ParentTextBoxContainer { get; }

        private void SetText(string value)
        {
            IParagraph baseParagraph = this.Paragraphs.First(p => p.Portions.Any());
            IEnumerable<IParagraph> removingParagraphs = this.Paragraphs.Where(p => p != baseParagraph);
            this.Paragraphs.Remove(removingParagraphs);

            baseParagraph.Text = value;
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

        public void ThrowIfRemoved() // TODO: make internal
        {
            ((Shape)this.ParentTextBoxContainer).ThrowIfRemoved();
        }
    }
}
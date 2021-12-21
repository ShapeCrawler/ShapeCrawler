using System;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using ShapeCrawler.Collections;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.AutoShapes
{
    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC — ShapeCrawler")]
    internal class SCTextBox : ITextBox
    {
        private readonly Lazy<string> text;

        public SCTextBox(OpenXmlCompositeElement txBodyCompositeElement, ITextBoxContainer parentTextBoxContainer)
        {
            this.text = new Lazy<string>(this.GetText);
            this.APTextBody = txBodyCompositeElement;
            this.Paragraphs = new ParagraphCollection(this.APTextBody, this);
            this.ParentTextBoxContainer = parentTextBoxContainer;
        }

        public IParagraphCollection Paragraphs { get; }

        public string Text
        {
            get => this.text.Value;
            set => this.SetText(value);
        }

        public AutofitType AutofitType => ParseAutofitType();

        internal ITextBoxContainer ParentTextBoxContainer { get; }

        internal OpenXmlCompositeElement APTextBody { get; }

        internal void ThrowIfRemoved()
        {
            this.ParentTextBoxContainer.ThrowIfRemoved();
        }

        private AutofitType ParseAutofitType()
        {
            var aBodyPr = this.APTextBody.GetFirstChild<A.BodyProperties>();
            if (aBodyPr!.GetFirstChild<A.NormalAutoFit>() != null)
            {
                return AutofitType.Shrink;
            }

            if (aBodyPr.GetFirstChild<A.ShapeAutoFit>() != null)
            {
                return AutofitType.Resize;
            }

            return AutofitType.None;
        }

        private void SetText(string value)
        {
            var baseParagraph = this.Paragraphs.First(p => p.Portions.Any());
            var removingParagraphs = this.Paragraphs.Where(p => p != baseParagraph);
            this.Paragraphs.Remove(removingParagraphs);

            if (this.AutofitType == AutofitType.Shrink)
            {
                var popularPortion = baseParagraph.Portions.GroupBy(p => p.Font.Size).OrderByDescending(x => x.Count()).First().First();
                var fontFamilyName = popularPortion.Font.Name;
                var fontSize = popularPortion.Font.Size;
                var font = new System.Drawing.Font(fontFamilyName, fontSize);
                var stringFormat = new StringFormat
                {
                    Trimming = StringTrimming.Word
                };
                var bm = new Bitmap(this.ParentTextBoxContainer.Shape.X, this.ParentTextBoxContainer.Shape.Y);
            }

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

    }
}
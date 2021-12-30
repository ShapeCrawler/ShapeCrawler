using System;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ShapeCrawler.Collections;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using Font = System.Drawing.Font;

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

        public AutofitType AutofitType => this.ParseAutofitType();

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

        private void SetText(string newText)
        {
            var baseParagraph = this.Paragraphs.First(p => p.Portions.Any());
            var removingParagraphs = this.Paragraphs.Where(p => p != baseParagraph);
            this.Paragraphs.Remove(removingParagraphs);

            if (this.AutofitType == AutofitType.Shrink)
            {
                var popularPortion = baseParagraph.Portions.GroupBy(p => p.Font.Size).OrderByDescending(x => x.Count()).First().First();
                var fontFamilyName = popularPortion.Font.Name;
                var fontSize = popularPortion.Font.Size;
                var stringFormat = new StringFormat { Trimming = StringTrimming.Word };
                var shape = this.ParentTextBoxContainer.Shape;
                var bm = new Bitmap(shape.Width, shape.Height);
                using var graphic = Graphics.FromImage(bm);
                var margin = 7;
                var rectangle = new Rectangle(margin, margin, shape.Width - 2 * margin, shape.Height - 2 * margin);
                var availSize = new SizeF(rectangle.Width, rectangle.Height);

                int charsFitted;
                do
                {
                    var font = new Font(fontFamilyName, fontSize);
                    graphic.MeasureString(newText, font, availSize, stringFormat, out charsFitted, out _);
                    fontSize--;
                }
                while (newText.Length != charsFitted);

                var paragraphInternal = (SCParagraph)baseParagraph;
                paragraphInternal.SetFontSize(fontSize);
            }

            baseParagraph.Text = newText;
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
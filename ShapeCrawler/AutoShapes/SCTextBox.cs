using System;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using ShapeCrawler.Collections;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using Font = System.Drawing.Font;

namespace ShapeCrawler.AutoShapes
{
    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC — ShapeCrawler")]
    internal class SCTextBox : ITextBox
    {
        private readonly ResettableLazy<string> text;
        private readonly ResettableLazy<ParagraphCollection> paragraphs;

        internal SCTextBox(ITextBoxContainer textBoxContainer, OpenXmlCompositeElement txBodyElement)
            : this(textBoxContainer)
        {
            this.APTextBody = txBodyElement;
        }
        
        internal SCTextBox(ITextBoxContainer textBoxContainer)
        {
            this.TextBoxContainer = textBoxContainer;
            this.text = new ResettableLazy<string>(this.GetText);
            this.paragraphs = new ResettableLazy<ParagraphCollection>(this.GetParagraphs);
        }



        public IParagraphCollection Paragraphs => this.paragraphs.Value;

        public string Text
        {
            get => this.text.Value;
            set => this.SetText(value);
        }

        public AutofitType AutofitType => this.ParseAutofitType();

        /// <summary>
        ///     Gets parent text box container.
        /// </summary>
        internal ITextBoxContainer TextBoxContainer { get; }

        internal OpenXmlCompositeElement? APTextBody { get; }

        internal void ThrowIfRemoved()
        {
            this.TextBoxContainer.ThrowIfRemoved();
        }

        private ParagraphCollection GetParagraphs()
        {
            return new ParagraphCollection(this);
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
                var popularPortion = baseParagraph.Portions.GroupBy(p => p.Font.Size).OrderByDescending(x => x.Count())
                    .First().First();
                var fontFamilyName = popularPortion.Font.Name;
                var fontSize = popularPortion.Font.Size;
                var stringFormat = new StringFormat { Trimming = StringTrimming.Word };
                var shape = this.TextBoxContainer.Shape;
                var bm = new Bitmap(shape.Width, shape.Height);
                using var graphic = Graphics.FromImage(bm);
                const int margin = 7;
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
            if (this.APTextBody == null)
            {
                return string.Empty;
            }
            
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
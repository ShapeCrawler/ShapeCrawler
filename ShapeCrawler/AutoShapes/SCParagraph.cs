using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Factories;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;


namespace ShapeCrawler.AutoShapes
{
    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC - ShapeCrawler")]
    internal class SCParagraph : IParagraph
    {
        private readonly Lazy<SCBullet> bullet;
        private readonly ResettableLazy<PortionCollection> portions;
        private TextAlignment? alignment;

        internal SCParagraph(A.Paragraph aParagraph, TextFrame textBox)
        {
            this.AParagraph = aParagraph;
            this.AParagraph.ParagraphProperties ??= new A.ParagraphProperties();
            this.Level = GetInnerLevel(aParagraph);
            this.bullet = new Lazy<SCBullet>(this.GetBullet);
            this.ParentTextBox = textBox;
            this.portions = new ResettableLazy<PortionCollection>(this.GetPortions);
        }

        public bool IsRemoved { get; set; }

        public string Text
        {
            get => this.GetText();
            set => this.SetText(value);
        }

        public IPortionCollection Portions => this.portions.Value;

        public SCBullet Bullet => this.bullet.Value;

        public TextAlignment Alignment
        {
            get => this.GetAlignment();
            set => this.UpdateAlignment(value);
        }

        internal TextFrame ParentTextBox { get; }

        internal A.Paragraph AParagraph { get; }

        internal int Level { get; }

        public void SetFontSize(int fontSize)
        {
            foreach (var portion in this.Portions)
            {
                portion.Font.Size = fontSize;
            }
        }

        public void AddPortion(string text)
        {
            this.ThrowIfRemoved();
            if (text == string.Empty)
            {
                return;
            }

            var basePortion = this.portions.Value.LastOrDefault();
            OpenXmlElement baseATextParent = null;
            OpenXmlElement lastARunOrABreak = null;
            if (basePortion == null)
            {
                baseATextParent = ARunInstance.CreateEmpty();
            }
            else
            {
                baseATextParent = basePortion.SDKAText.Parent!;
                lastARunOrABreak = this.AParagraph.Last(p => p is A.Run or A.Break);
            }

            // add break if last element is not A.Break && text ends with newLine
            if (lastARunOrABreak is not A.Break && this.Text.EndsWith(Environment.NewLine, StringComparison.Ordinal))
            {
                AddBreak(ref lastARunOrABreak);
            }

            string[] textLines = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            if (basePortion?.Text == string.Empty)
            {
                basePortion.Text = textLines[0];
            }
            else
            {
                AddText(ref lastARunOrABreak, baseATextParent, textLines[0], this.AParagraph);
            }

            for (int i = 1; i < textLines.Length; i++)
            {
                AddBreak(ref lastARunOrABreak);

                if (textLines[i] != string.Empty)
                {
                    AddText(ref lastARunOrABreak, baseATextParent, textLines[i], this.AParagraph);
                }
            }

            this.portions.Reset();
        }

        internal void ThrowIfRemoved()
        {
            if (this.IsRemoved)
            {
                throw new ElementIsRemovedException("Paragraph was removed.");
            }

            this.ParentTextBox.ThrowIfRemoved();
        }

        #region Private Methods

        private static void AddBreak(ref OpenXmlElement lastElement)
        {
            lastElement = lastElement.InsertAfterSelf(new A.Break());
        }

        private static void AddText(ref OpenXmlElement lastElement, OpenXmlElement basePortionElement, string text,
            A.Paragraph aParagraph)
        {
            var newARun = (A.Run)basePortionElement.CloneNode(true);
            newARun.Text.Text = text;
            if (lastElement == null)
            {
                aParagraph.InsertAt(newARun, 0);
            }
            else
            {
                lastElement = lastElement.InsertAfterSelf(newARun);
            }
        }

        private static int GetInnerLevel(A.Paragraph aParagraph)
        {
            // XML-paragraph enumeration started from zero. Null is also zero
            Int32Value xmlParagraphLvl = aParagraph.ParagraphProperties?.Level ?? 0;
            int paragraphLvl = ++xmlParagraphLvl;

            return paragraphLvl;
        }

        private SCBullet GetBullet()
        {
            return new SCBullet(this.AParagraph.ParagraphProperties);
        }

        private string GetText()
        {
            if (this.Portions.Count == 0)
            {
                return string.Empty;
            }

            return this.Portions.Select(portion => portion.Text).Aggregate((result, next) => result + next);
        }

        private void SetText(string newText)
        {
            this.ThrowIfRemoved();

            // To set a paragraph text we use a single portion which is the first paragraph portion.
            // Rest of the portions are deleted from the paragraph.
            var removingPortions = this.Portions.Skip(1).ToList();
            this.Portions.Remove(removingPortions);
            var basePortion = (SCPortion)this.portions.Value.Single();

            basePortion.Text = String.Empty;
            this.AddPortion(newText);
        }

        private PortionCollection GetPortions()
        {
            return new PortionCollection(this.AParagraph, this);
        }

        private void UpdateAlignment(TextAlignment alignmentValue)
        {
            if (this.ParentTextBox.TextFrameContainer.Placeholder != null)
            {
                throw new PlaceholderCannotBeChangedException();
            }

            A.TextAlignmentTypeValues sdkAlignmentValue = alignmentValue switch
            {
                TextAlignment.Left => A.TextAlignmentTypeValues.Left,
                TextAlignment.Center => A.TextAlignmentTypeValues.Center,
                TextAlignment.Right => A.TextAlignmentTypeValues.Right,
                TextAlignment.Justify => A.TextAlignmentTypeValues.Justified,
                _ => throw new ArgumentOutOfRangeException(nameof(alignmentValue))
            };

            if (this.AParagraph.ParagraphProperties == null)
            {
                this.AParagraph.ParagraphProperties = new A.ParagraphProperties
                {
                    Alignment = new EnumValue<A.TextAlignmentTypeValues>(sdkAlignmentValue)
                };
            }
            else
            {
                this.AParagraph.ParagraphProperties.Alignment = new EnumValue<A.TextAlignmentTypeValues>(sdkAlignmentValue);
            }

            this.alignment = alignmentValue;
        }

        private TextAlignment GetAlignment()
        {
            if (this.alignment.HasValue)
            {
                return this.alignment.Value;
            }

            var placeholder = this.ParentTextBox.TextFrameContainer.Placeholder;
            if (placeholder is { Type: PlaceholderType.Title })
            {
                this.alignment = TextAlignment.Left;
                return this.alignment.Value;
            }

            if (placeholder is { Type: PlaceholderType.CenteredTitle })
            {
                this.alignment = TextAlignment.Center;
                return this.alignment.Value;
            }

            var algnAttribute = this.AParagraph.ParagraphProperties?.Alignment!;
            if (algnAttribute == null)
            {
                return TextAlignment.Left;
            }

            this.alignment = algnAttribute.Value switch
            {
                A.TextAlignmentTypeValues.Center => TextAlignment.Center,
                A.TextAlignmentTypeValues.Right => TextAlignment.Right,
                A.TextAlignmentTypeValues.Justified => TextAlignment.Justify,
                _ => TextAlignment.Left
            };

            return this.alignment.Value;
        }

        #endregion Private Methods
    }
}
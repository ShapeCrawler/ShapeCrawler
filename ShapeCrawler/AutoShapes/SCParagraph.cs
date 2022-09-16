using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shared;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
// ReSharper disable SuggestVarOrType_SimpleTypes
// ReSharper disable SuggestVarOrType_BuiltInTypes
namespace ShapeCrawler
{
    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC - ShapeCrawler")]
    internal class SCParagraph : IParagraph
    {
        private readonly Lazy<Bullet> bullet;
        private readonly ResettableLazy<PortionCollection> portions;
        private TextAlignment? alignment;

        internal SCParagraph(A.Paragraph aParagraph, SCTextBox textBox)
        {
            this.AParagraph = aParagraph;
            this.Level = GetInnerLevel(aParagraph);
            this.bullet = new Lazy<Bullet>(this.GetBullet);
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

        public Bullet Bullet => this.bullet.Value;

        public TextAlignment Alignment
        {
            get => this.GetAlignment();
            set => this.UpdateAlignment(value);
        }

        internal SCTextBox ParentTextBox { get; }

        internal A.Paragraph AParagraph { get; }

        internal int Level { get; }

        public void SetFontSize(int fontSize)
        {
            foreach (var portion in this.Portions)
            {
                portion.Font.Size = fontSize;
            }
        }

        internal void ThrowIfRemoved()
        {
            if (this.IsRemoved)
            {
                throw new ElementIsRemovedException("Paragraph was removed.");
            }
            else
            {
                this.ParentTextBox.ThrowIfRemoved();
            }
        }

        #region Private Methods

        private static int GetInnerLevel(A.Paragraph aParagraph)
        {
            // XML-paragraph enumeration started from zero. Null is also zero
            Int32Value xmlParagraphLvl = aParagraph.ParagraphProperties?.Level ?? 0;
            int paragraphLvl = ++xmlParagraphLvl;

            return paragraphLvl;
        }

        private Bullet GetBullet()
        {
            return new Bullet(this.AParagraph.ParagraphProperties);
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
            AddPortion(newText);
        }

        public void AddPortion(string sourceText)
        {
            void addBreak(ref OpenXmlElement lastElement)
            {
                lastElement = lastElement.InsertAfterSelf(new A.Break());
            }

            void addText(ref OpenXmlElement lastElement, OpenXmlElement basePortionElement, string text)
            {
                var newARun = (A.Run)basePortionElement.CloneNode(true);
                newARun.Text.Text = text;
                lastElement = lastElement.InsertAfterSelf(newARun);
            }

            this.ThrowIfRemoved();
            if (sourceText == String.Empty)
            {
                this.portions.Reset();
                return;
            }

            var basePortion = (SCPortion)this.portions.Value.Last();
            var basePortionElement = basePortion.SDKAText.Parent;
            var lastElement = this.AParagraph.Where(p => p is A.Run || p is A.Break).Last();
            
            // add break if last element is not A.Break && text ends with newLine
            if (lastElement is not A.Break && this.Text.EndsWith(Environment.NewLine, StringComparison.Ordinal))
            {
                addBreak(ref lastElement);
            }

            string[] textLines = sourceText.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            if (basePortion.Text == String.Empty)
            {
                basePortion.Text = textLines[0];
            }
            else
            {
                addText(ref lastElement, basePortionElement, textLines[0]);
            }
            
            for (int i = 1; i < textLines.Length; i++)
            {
                addBreak(ref lastElement);

                if (textLines[i] != string.Empty)
                {
                    addText(ref lastElement, basePortionElement, textLines[i]);
                }
            }

            this.portions.Reset();
        }

        private PortionCollection GetPortions()
        {
            return new PortionCollection(this.AParagraph, this);
        }
        
        private void UpdateAlignment(TextAlignment alignmentValue)
        {
            if (this.ParentTextBox.TextBoxContainer.Placeholder != null)
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

            var placeholder = this.ParentTextBox.TextBoxContainer.Placeholder;
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
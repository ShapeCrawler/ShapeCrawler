using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shared;
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

        internal SCParagraph(A.Paragraph aParagraph, SCTextBox textBox)
        {
            this.AParagraph = aParagraph;
            this.Level = GetInnerLevel(aParagraph);
            this.bullet = new Lazy<Bullet>(this.GetBullet);
            this.ParentTextBox = textBox;
            this.portions = new ResettableLazy<PortionCollection>(() => new PortionCollection(this.AParagraph, this));
        }

        #region Public Properties

        public bool IsRemoved { get; set; }

        public string Text
        {
            get => this.GetText();
            set => this.SetText(value);
        }

        public IPortionCollection Portions => this.portions.Value;

        public Bullet Bullet => this.bullet.Value;

        #endregion Public Properties

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

            if (newText == string.Empty)
            {
                basePortion.Text = string.Empty;
                return;
            }

            string[] textLines = newText.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            basePortion.Text = textLines[0];
            OpenXmlElement lastInsertedARunOrLineBreak = basePortion.AText.Parent;
            for (int i = 1; i < textLines.Length; i++)
            {
                lastInsertedARunOrLineBreak = lastInsertedARunOrLineBreak.InsertAfterSelf(new A.Break());
                A.Run newARun = (A.Run)basePortion.AText.Parent.CloneNode(true);
                newARun.Text.Text = textLines[i];
                lastInsertedARunOrLineBreak = lastInsertedARunOrLineBreak.InsertAfterSelf(newARun);
            }

            if (newText.EndsWith(Environment.NewLine, StringComparison.Ordinal))
            {
                lastInsertedARunOrLineBreak.InsertAfterSelf(new A.Break());
            }

            this.portions.Reset();
        }

        #endregion Private Methods
    }
}
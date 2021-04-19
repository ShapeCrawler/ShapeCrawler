using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Collections;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
// ReSharper disable SuggestVarOrType_SimpleTypes
// ReSharper disable SuggestVarOrType_BuiltInTypes

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a text paragraph.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC - ShapeCrawler")]
    internal class SCParagraph : IParagraph
    {
        private readonly Lazy<Bullet> bullet;
        private readonly ResettableLazy<PortionCollection> portions;

        /// <summary>
        ///     Initializes a new instance of the <see cref="SCParagraph"/> class.
        /// </summary>
        internal SCParagraph(A.Paragraph aParagraph, SCTextBox textBox)
        {
            this.AParagraph = aParagraph;
            this.Level = GetInnerLevel(aParagraph);
            this.bullet = new Lazy<Bullet>(this.GetBullet);
            this.TextBox = textBox;
            this.portions = new ResettableLazy<PortionCollection>(() => new PortionCollection(this.AParagraph, this));
        }

        /// <summary>
        ///     Gets or sets the the plain text of a paragraph.
        /// </summary>
        public string Text
        {
            get => this.GetText();
            set => this.SetText(value);
        }

        /// <summary>
        ///     Gets collection of paragraph portions. Returns <c>NULL</c> if paragraph is empty.
        /// </summary>
        public IPortionCollection Portions => this.portions.Value;

        /// <summary>
        ///     Gets paragraph bullet. Returns <c>NULL</c> if bullet does not exist.
        /// </summary>
        public Bullet Bullet => bullet.Value;

        internal SCTextBox TextBox { get; }

        internal A.Paragraph AParagraph { get; }

        internal int Level { get; }

       

        #region Private Methods

        private Bullet GetBullet()
        {
            return new Bullet(AParagraph.ParagraphProperties);
        }

        private static int GetInnerLevel(A.Paragraph aParagraph)
        {
            // XML-paragraph enumeration started from zero. Null is also zero
            Int32Value xmlParagraphLvl = aParagraph.ParagraphProperties?.Level ?? 0;
            int paragraphLvl = ++xmlParagraphLvl;

            return paragraphLvl;
        }

        private string GetText()
        {
            if (this.Portions.Count == 0)
            {
                return string.Empty;
            }

            return Portions.Select(portion => portion.Text).Aggregate((result, next) => result + next);
        }

        private void SetText(string newText)
        {
            // To set a paragraph text we use a single portion which is the first paragraph portion.
            // Rest of the portions are deleted from the paragraph.
            Portions.Remove(Portions.Skip(1).ToList());
            IPortion basePortion = Portions.Single();
            if (newText == string.Empty)
            {
                basePortion.Text = string.Empty;
                return;
            }

            string[] textLines = newText.Split(new[] {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries);
            basePortion.Text = textLines[0];
            OpenXmlElement lastInsertedARunOrLineBreak = ((Portion) basePortion).AText.Parent;
            for (int i = 1; i < textLines.Length; i++)
            {
                lastInsertedARunOrLineBreak = lastInsertedARunOrLineBreak.InsertAfterSelf(new A.Break());
                A.Run newARun = ((Portion) basePortion).GetARunCopy();
                newARun.Text.Text = textLines[i];
                lastInsertedARunOrLineBreak = lastInsertedARunOrLineBreak.InsertAfterSelf(newARun);
            }

            if (newText.EndsWith(Environment.NewLine, StringComparison.Ordinal))
            {
                lastInsertedARunOrLineBreak.InsertAfterSelf(new A.Break());
            }

            portions.Reset();
        }

        #endregion Private Methods
    }
}
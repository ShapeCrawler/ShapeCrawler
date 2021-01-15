using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Collections;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable PossibleMultipleEnumeration
// ReSharper disable SuggestVarOrType_SimpleTypes
// ReSharper disable SuggestVarOrType_BuiltInTypes

namespace ShapeCrawler.Models.TextShape
{
    /// <summary>
    /// Represents a text paragraph.
    /// </summary>
    [SuppressMessage("ReSharper", "SuggestVarOrType_Elsewhere")]
    public class Paragraph
    {
        #region Fields

        private readonly A.Paragraph _textParagraph;
        private readonly Lazy<Bullet> _bullet;
        private readonly ResettableLazy<PortionCollection> _portions;

        #endregion Fields

        internal TextSc TextFrame { get; }
        internal int Level { get; }

        #region Public Properties

        /// <summary>
        /// Gets or sets the the plain text of a paragraph.
        /// </summary>
        public string Text
        {
            get => GetText();
            set => SetText(value);
        }

        /// <summary>
        /// Gets collection of paragraph text portions.
        /// </summary>
        public PortionCollection Portions => _portions.Value;

        /// <summary>
        /// Gets paragraph bullet. Returns null if bullet does not exist.
        /// </summary>
        public Bullet Bullet => _bullet.Value;

        #endregion Public Properties

        #region Constructors

        /// <summary>
        /// Initializes an instance of the <see cref="Paragraph"/> class.
        /// </summary>
        public Paragraph(ShapeContext spContext, A.Paragraph aParagraph, TextSc textFrame) //TODO: Replace constructor initialization on static .Create()
        {
            _textParagraph = aParagraph;
            Level = GetInnerLevel(aParagraph);
            _bullet = new Lazy<Bullet>(GetBullet);
            TextFrame = textFrame;
            _portions = new ResettableLazy<PortionCollection>(() => PortionCollection.Create(_textParagraph, spContext, this));
        }

        private Bullet GetBullet()
        {
            return new Bullet(_textParagraph.ParagraphProperties);
        }

        #endregion Constructors

        #region Private Methods

        private static int GetInnerLevel(A.Paragraph aParagraph)
        {
            // XML-paragraph enumeration started from zero. Null is also zero
            Int32Value sdkParagraphLvl = aParagraph.ParagraphProperties?.Level ?? 0;
            int paragraphLvl = ++sdkParagraphLvl;

            return paragraphLvl;
        }

        private string GetText()
        {
            return Portions.Select(p => p.Text).Aggregate((result, next) => result + next);
        }

        private void SetText(string newText)
        {
            // TODO: Improve deleting performance, for example by adding a new method RemoveAllExceptFirst

            // To set a paragraph text we use a single portion which is the first paragraph portion.
            // Rest of the portions are deleted from the paragraph.
            Portions.RemoveRange(Portions.Skip(1).ToList());
            Portion basePortion = Portions.Single();
            string[] textLines = newText.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            basePortion.Text = textLines[0];
            OpenXmlElement lastInsertedARunOrLineBreak = basePortion.AText.Parent;
            for (int i = 1; i < textLines.Length; i++)
            {
                lastInsertedARunOrLineBreak = lastInsertedARunOrLineBreak.InsertAfterSelf(new A.Break());
                A.Run newARun = basePortion.GetARunCopy();
                newARun.Text.Text = textLines[i];
                lastInsertedARunOrLineBreak = lastInsertedARunOrLineBreak.InsertAfterSelf(newARun);
            }

            if (newText.EndsWith(Environment.NewLine, StringComparison.Ordinal))
            {
                lastInsertedARunOrLineBreak.InsertAfterSelf(new A.Break());
            }

            _portions.Reset();
        }

        #endregion Private Methods
    }
}

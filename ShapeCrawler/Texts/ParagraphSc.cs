﻿using System;
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

namespace ShapeCrawler.Texts
{
    /// <summary>
    /// Represents a text paragraph.
    /// </summary>
    [SuppressMessage("ReSharper", "SuggestVarOrType_Elsewhere")]
    public class ParagraphSc
    {
        #region Fields

        private readonly Lazy<Bullet> _bullet;
        private readonly ResettableLazy<PortionCollection> _portions;

        internal TextBoxSc TextBox { get; }
        internal A.Paragraph AParagraph { get; }
        internal int Level { get; }

        #endregion Fields

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
        /// Initializes an instance of the <see cref="ParagraphSc"/> class.
        /// </summary>
        // TODO: Replace constructor initialization on static .Create()
        internal ParagraphSc(A.Paragraph aParagraph, TextBoxSc textBox)
        {
            AParagraph = aParagraph;
            Level = GetInnerLevel(aParagraph);
            _bullet = new Lazy<Bullet>(GetBullet);
            TextBox = textBox;
            _portions = new ResettableLazy<PortionCollection>(() => PortionCollection.Create(AParagraph, this));
        }

        #endregion Constructors

        #region Private Methods

        private Bullet GetBullet()
        {
            return new Bullet(AParagraph.ParagraphProperties);
        }

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
            // TODO: Add RemoveRange API to remove all portion except first

            // To set a paragraph text we use a single portion which is the first paragraph portion.
            // Rest of the portions are deleted from the paragraph.
            Portions.Remove(Portions.Skip(1).ToList());
            Portion basePortion = Portions.Single();
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

        internal void Remove()
        {
            AParagraph.Remove();
        }
    }
}

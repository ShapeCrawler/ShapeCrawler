﻿using System;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes
{
    /// <summary>
    ///     Represents a text paragraph portion.
    /// </summary>
    public class Portion // TODO: add interface
    {
        private readonly ResettableLazy<FontSc> _font;
        internal readonly A.Text AText;

        #region Constructors

        internal Portion(A.Text aText, ParagraphSc paragraph)
        {
            AText = aText;
            Paragraph = paragraph;
            _font = new ResettableLazy<FontSc>(GetFont);
        }

        #endregion Constructors

        internal ParagraphSc Paragraph { get; }

        internal A.Run GetARunCopy()
        {
            return (A.Run) AText.Parent.CloneNode(true);
        }

        #region Public Properties

        /// <summary>
        ///     Gets or sets paragraph portion text.
        /// </summary>
        public string Text
        {
            get => GetText();
            set => SetText(value);
        }

        /// <summary>
        ///     Gets font.
        /// </summary>
        public FontSc Font => _font.Value;

        /// <summary>
        ///     Removes portion from the paragraph.
        /// </summary>
        public void Remove()
        {
            Paragraph.Portions.Remove(this);
        }

        #endregion Public Properties

        #region Private Methods

        private FontSc GetFont()
        {
            return new FontSc(AText, this);
        }

        private string GetText()
        {
            string portionText = AText.Text;
            if (AText.Parent.NextSibling<A.Break>() != null)
            {
                portionText += Environment.NewLine;
            }

            return portionText;
        }

        private void SetText(string text)
        {
            AText.Text = text;
        }

        #endregion Private Methods
    }
}
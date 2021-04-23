using System;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace

namespace ShapeCrawler
{
    /// <inheritdoc cref="IPortion"/>
    internal class Portion : IPortion // TODO: make internal
    {
        private readonly ResettableLazy<SCFont> font;

        /// <summary>
        ///     Initializes a new instance of the <see cref="Portion"/> class.
        /// </summary>
        internal Portion(A.Text aText, SCParagraph paragraph)
        {
            this.AText = aText;
            this.ParentParagraph = paragraph;
            this.font = new ResettableLazy<SCFont>(this.GetFont);
        }

        #region Public Properties

        /// <inheritdoc/>
        public string Text
        {
            get => this.GetText();

            set
            {
                this.ThrowIfRemoved();
                this.SetText(value);
            }
        }

        /// <inheritdoc/>
        public IFont Font => this.font.Value;

        #endregion Public Properties

        internal bool IsRemoved { get; set; }

        internal SCParagraph ParentParagraph { get; }

        internal A.Text AText { get; }

        private void ThrowIfRemoved()
        {
            if (this.IsRemoved)
            {
                throw new ElementIsRemovedException("Paragraph portion was removed.");
            }
            else
            {
                this.ParentParagraph.ThrowIfRemoved();
            }
        }

        private SCFont GetFont()
        {
            return new SCFont(this.AText, this);
        }

        private string GetText()
        {
            string portionText = this.AText.Text;
            if (this.AText.Parent.NextSibling<A.Break>() != null)
            {
                portionText += Environment.NewLine;
            }

            return portionText;
        }

        private void SetText(string text)
        {
            this.AText.Text = text;
        }
    }
}
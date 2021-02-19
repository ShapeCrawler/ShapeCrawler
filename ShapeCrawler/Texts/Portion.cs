using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Texts
{
    /// <summary>
    /// Represents a text paragraph portion.
    /// </summary>
    public class Portion
    {
        private readonly ResettableLazy<FontSc> _font;

        internal ParagraphSc Paragraph { get; }
        internal readonly A.Text AText;

        #region Public Properties

        /// <summary>
        /// Gets or sets paragraph portion text.
        /// </summary>
        public string Text
        {
            get => GetText();
            set => SetText(value);
        }

        /// <summary>
        /// Gets font.
        /// </summary>
        public FontSc Font => _font.Value;
        
        /// <summary>
        /// Removes portion from the paragraph.
        /// </summary>
        public void Remove()
        {
            Paragraph.Portions.Remove(this);
        }

        #endregion Public Properties

        #region Constructors

        internal Portion(A.Text aText, ParagraphSc paragraph)
        {
            AText = aText;
            Paragraph = paragraph;
            _font = new ResettableLazy<FontSc>(GetFont);
        }

        #endregion Constructors

        #region Private Methods

        private FontSc GetFont()
        {
            int fontSize = GetFontSize();
            return new FontSc(AText, fontSize, this);
        }

        private int GetFontSize()
        {
            Int32Value aRunPropertiesSize = AText.Parent.GetFirstChild<A.RunProperties>()?.FontSize;
            if (aRunPropertiesSize != null)
            {
                return aRunPropertiesSize.Value;
            }

            // If element is placeholder, tries to get from placeholder data
            P.Shape shapeTreeSource = (P.Shape)Paragraph.TextBox.AutoShape.ShapeTreeSource;
            if (shapeTreeSource.IsPlaceholder())
            {
                int? prFontHeight = Paragraph.TextBox.ShapeContext.PlaceholderFontService.GetFontSizeByParagraphLvl(shapeTreeSource, Paragraph.Level);
                if (prFontHeight != null)
                {
                    return (int)prFontHeight;
                }
            }

            PresentationData presentationData = Paragraph.TextBox.AutoShape.Slide.Presentation.PresentationData;
            if (presentationData.LlvFontHeights.ContainsKey(Paragraph.Level))
            {
                return presentationData.LlvFontHeights[Paragraph.Level];
            }

            var exist = Paragraph.TextBox.ShapeContext.TryGetFromMasterOtherStyle(Paragraph.Level, out int fh);
            if (exist)
            {
                return fh;
            }

            return FormatConstants.DefaultFontSize;
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

        internal A.Run GetARunCopy()
        {
            return (A.Run) AText.Parent.CloneNode(true);
        }
    }
}
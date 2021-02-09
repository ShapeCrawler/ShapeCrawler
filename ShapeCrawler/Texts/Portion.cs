using System;
using System.Diagnostics.CodeAnalysis;
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
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    [SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
    public class Portion
    {
        private readonly ResettableLazy<FontSc> _font;
        private readonly ShapeContext _shapeContext;

        #region Internal Properties

        internal ParagraphSc Paragraph { get; }
        
        internal readonly A.Text AText;

        #endregion Internal Properties

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

        internal Portion(A.Text aText, ParagraphSc paragraph, ShapeContext shapeContext)
        {
            AText = aText;
            Paragraph = paragraph;
            _shapeContext = shapeContext;
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
            P.Shape shapeTreeSource = (P.Shape)Paragraph.TextBox.Shape.ShapeTreeSource;
            if (shapeTreeSource.IsPlaceholder())
            {
                int? prFontHeight = _shapeContext.PlaceholderFontService.GetFontSizeByParagraphLvl(shapeTreeSource, Paragraph.Level);
                if (prFontHeight != null)
                {
                    return (int)prFontHeight;
                }
            }

            PresentationData presentationData = Paragraph.TextBox.Shape.Slide.Presentation.PresentationData;
            if (presentationData.LlvFontHeights.ContainsKey(Paragraph.Level))
            {
                return presentationData.LlvFontHeights[Paragraph.Level];
            }

            var exist = _shapeContext.TryGetFromMasterOtherStyle(Paragraph.Level, out int fh);
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
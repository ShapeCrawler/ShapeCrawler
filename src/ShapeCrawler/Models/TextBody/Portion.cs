using System.Linq;
using ShapeCrawler.Collections;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Models.TextBody
{
    /// <summary>
    /// Represents a portion of text inside a text paragraph.
    /// </summary>
    public class Portion
    {
        private readonly A.Text _aText;
        private readonly Paragraph _paragraph;

        #region Properties

        /// <summary>
        /// Gets font height in EMUs.
        /// </summary>
        public int FontHeight { get; }

        /// <summary>
        /// Gets or sets text.
        /// </summary>
        public string Text
        {
            get => _aText.Text;
            set => _aText.Text = value;
        }

        public string FontName
        {
            get => ParseFontName();
            set
            {

            }
        }

        private string ParseFontName()
        {
            var rPr = _aText.Parent.GetFirstChild<A.RunProperties>();
            var latinFont = rPr?.GetFirstChild<A.LatinFont>();
            if (latinFont == null)
            {
                // Gets font from theme
                latinFont = _aText.Ancestors<P.Slide>().Single()
                    .SlidePart.SlideLayoutPart.SlideMasterPart
                    .ThemePart.Theme.ThemeElements.FontScheme.MinorFont
                    .LatinFont;
            }

            return latinFont.Typeface;
        }

        /// <summary>
        /// Removes the portion from paragraph.
        /// </summary>
        public void Remove()
        {
            _paragraph.Portions.Remove(this);
        }
        
        #endregion Properties

        #region Constructors

        public Portion(A.Text aText, int fontHeight, Paragraph paragraph)
        {
            Check.IsPositive(fontHeight, nameof(fontHeight));
            FontHeight = fontHeight;
            _aText = aText;
            _paragraph = paragraph;
        }

        #endregion Constructors
    }
}
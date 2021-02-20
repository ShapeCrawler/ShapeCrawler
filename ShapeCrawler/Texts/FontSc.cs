using System.Linq;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Texts
{
    public class FontSc
    {
        private readonly A.Text _aText;
        private readonly ResettableLazy<A.LatinFont> _latinFont;
        private readonly Portion _portion;
        private readonly int _size;

        #region Constructors

        internal FontSc(A.Text aText, int fontSize, Portion portion)
        {
            _aText = aText;
            _size = fontSize;
            _latinFont = new ResettableLazy<A.LatinFont>(ParseLatinFont);
            _portion = portion;
        }

        #endregion Constructors

        #region Public Properties

        /// <summary>
        ///     Gets font name.
        /// </summary>
        public string Name
        {
            get => ParseFontName();
            set => SetFontName(value);
        }

        /// <summary>
        ///     Gets or sets font size in EMUs.
        /// </summary>
        public int Size
        {
            get => _size;
            set => SetFontSize(value);
        }

        /// <summary>
        ///     Gets value indicating whether font size can be changed.
        /// </summary>
        public bool SizeCanBeChanged()
        {
            var runPr = _aText.Parent.GetFirstChild<A.RunProperties>();
            return runPr != null;
        }

        #endregion Public Properties

        #region Private Methods

        private string ParseFontName()
        {
            return _latinFont.Value.Typeface;
        }

        private void SetFontName(string fontName)
        {
            if (_portion.Paragraph.TextBox.AutoShape.Placeholder != null)
            {
                throw new PlaceholderCannotBeChangedException();
            }

            A.LatinFont latinFont = _latinFont.Value;
            latinFont.Typeface = fontName;
            _latinFont.Reset();
        }

        private void SetFontSize(int newFontSize)
        {
            var runPr = _aText.Parent.GetFirstChild<A.RunProperties>();
            if (runPr == null)
            {
                const string errorMsg =
                    "The property value cannot be changed on the Slide level since it belongs to Slide Master. " +
                    "Hence, you should change it on Slide Master level. " +
                    "Note: you can check whether the property can be changed via {property_name}CanBeChanged method.";
                throw new SlideMasterPropertyCannotBeChanged(errorMsg);
            }

            runPr.FontSize = newFontSize;
        }

        private A.LatinFont ParseLatinFont()
        {
            var runPr = _aText.Parent.GetFirstChild<A.RunProperties>();
            var latinFont = runPr?.GetFirstChild<A.LatinFont>();
            if (latinFont == null)
            {
                // Gets font from theme
                latinFont = _aText.Ancestors<P.Slide>().First()
                    .SlidePart.SlideLayoutPart.SlideMasterPart
                    .ThemePart.Theme.ThemeElements.FontScheme.MinorFont
                    .LatinFont;
            }

            return latinFont;
        }

        #endregion Private Methods
    }
}
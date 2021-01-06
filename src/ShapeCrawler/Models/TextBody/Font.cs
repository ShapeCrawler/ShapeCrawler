using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Models.TextBody
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    public class Font
    {
        private readonly A.Text _aText;
        private readonly Portion _portion;
        private readonly ResettableLazy<A.LatinFont> _latinFont;

        public Font(A.Text aText, int fontSize, Portion portion)
        {
            _aText = aText;
            Size = fontSize;
            _latinFont = new ResettableLazy<A.LatinFont>(ParseLatinFont);
            _portion = portion;
        }

        public string Name
        {
            get => ParseFontName();
            set => SetFontName(value);
        }

        public int Size { get; }

        private string ParseFontName()
        {
            return _latinFont.Value.Typeface;
        }

        private void SetFontName(string fontName)
        {
            if (_portion.Paragraph.TextFrame.Shape.Placeholder != null)
            {
                throw new PlaceholderCannotBeChangedException();
            }
            A.LatinFont latinFont = _latinFont.Value;
            latinFont.Typeface = fontName;
            _latinFont.Reset();
        }

        private A.LatinFont ParseLatinFont()
        {
            var rPr = _aText.Parent.GetFirstChild<A.RunProperties>();
            var latinFont = rPr?.GetFirstChild<A.LatinFont>();
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
    }
}
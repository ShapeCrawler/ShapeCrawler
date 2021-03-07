using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes
{
    public class FontSc
    {
        private readonly A.Text _aText;
        private readonly ResettableLazy<A.LatinFont> _latinFont;
        private readonly Portion _portion;
        private readonly ResettableLazy<int> _size;

        #region Constructors

        internal FontSc(A.Text aText, Portion portion)
        {
            _aText = aText;
            _size = new ResettableLazy<int>(GetSize);
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
            get => _size.Value;
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
            const string majorLatinFont = "+mj-lt";
            if (_latinFont.Value.Typeface == majorLatinFont)
            {
                return _portion.Paragraph.TextBox.AutoShape.ThemePart.Theme.ThemeElements.FontScheme.MajorFont.LatinFont
                    .Typeface;
            }

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
            A.RunProperties aRunPr = _aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunPr == null)
            {
                const string errorMsg =
                    "The property value cannot be changed on the Slide level since it belongs to Slide Master. " +
                    "Hence, you should change it on Slide Master level. " +
                    "Note: you can check whether the property can be changed via {property_name}CanBeChanged method.";
                throw new SlideMasterPropertyCannotBeChanged(errorMsg);
            }

            aRunPr.FontSize = newFontSize;
        }

        private A.LatinFont ParseLatinFont()
        {
            A.RunProperties aRunPr = _aText.Parent.GetFirstChild<A.RunProperties>();
            A.LatinFont aLatinFont = aRunPr?.GetFirstChild<A.LatinFont>();
            if (aLatinFont == null)
            {
                aLatinFont = _portion.Paragraph.TextBox.AutoShape.ThemePart.Theme.ThemeElements.FontScheme.MinorFont.LatinFont;
            }

            return aLatinFont;
        }

        private int GetSize()
        {
            Int32Value aRunPrFontSize = _portion.AText.Parent.GetFirstChild<A.RunProperties>()?.FontSize;
            if (aRunPrFontSize != null)
            {
                return aRunPrFontSize.Value;
            }

            ShapeContext shapeContext = _portion.Paragraph.TextBox.ShapeContext;
            Shape autoShape = _portion.Paragraph.TextBox.AutoShape;
            int paragraphLvl = _portion.Paragraph.Level;

            // NEW
            // Try get font size from placeholder
            if (autoShape is not MasterAutoShape && autoShape.Placeholder != null)
            {
                Placeholder placeholder = (Placeholder) autoShape.Placeholder;
                IAutoShapeInternal placeholderAutoShape = (IAutoShapeInternal) placeholder.Shape;
                if (placeholderAutoShape.TryGetFontSize(paragraphLvl, out int fontSize))
                {
                    return fontSize;
                }
            }

            if (autoShape.Placeholder != null)
            {
                // From Slide Master body
                if (autoShape.SlideMaster.TryGetFontSizeFromBody(paragraphLvl, out int fontSizeBody))
                {
                    return fontSizeBody;
                }

                // From Slide Master other
                if (autoShape.SlideMaster.TryGetFontSizeFromOther(paragraphLvl, out int fontSizeOther))
                {
                    return fontSizeOther;
                }
            }

            // From presentation level
            PresentationData presentationData = autoShape.Presentation.PresentationData;
            if (presentationData.LlvToFontData.TryGetValue(paragraphLvl, out FontData fontData))
            {
                if (fontData.FontSize != null)
                {
                    return fontData.FontSize;
                }
            }

            

            return FormatConstants.DefaultFontSize;
        }

        #endregion Private Methods
    }

    internal interface IAutoShapeInternal
    {
        bool TryGetFontSize(int paragraphLvl, out int i);
    }
}
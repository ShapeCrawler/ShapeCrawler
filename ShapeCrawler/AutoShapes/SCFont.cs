using System.Diagnostics.CodeAnalysis;
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
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    internal class SCFont : IFont
    {
        private readonly A.Text _aText;
        private readonly ResettableLazy<A.LatinFont> _latinFont;
        private readonly int _paragraphLvl;
        private readonly Portion _portion;
        private readonly ResettableLazy<int> _size;

        #region Constructors

        internal SCFont(A.Text aText, Portion portion)
        {
            _aText = aText;
            _size = new ResettableLazy<int>(GetSize);
            _latinFont = new ResettableLazy<A.LatinFont>(GetALatinFont);
            _portion = portion;
            _paragraphLvl = _portion.Paragraph.Level;
        }

        #endregion Constructors

        #region Public Properties

        /// <summary>
        ///     Gets font name.
        /// </summary>
        public string Name
        {
            get => GetName();
            set => SetName(value);
        }

        /// <summary>
        ///     Gets or sets font size in EMUs.
        /// </summary>
        public int Size
        {
            get => _size.Value;
            set => SetFontSize(value);
        }

        public bool IsBold
        {
            get => GetBoldFlag();
            set => SetBoldFlag(value);
        }

        public bool IsItalic
        {
            get => GetItalicFlag();
            set => SetItalicFlag(value);
        }

        public string Color
        {
            get => GetColorHex();
            set => SetColorHex(value);
        }

        private void SetColorHex(string value)
        {
        }

        private string GetColorHex()
        {
            // Try get color from PORTION level
            A.SolidFill aSolidFill = _portion.AText.PreviousSibling<A.RunProperties>()?.GetFirstChild<A.SolidFill>();
            if (aSolidFill != null)
            {
                A.RgbColorModelHex hexModel = aSolidFill.RgbColorModelHex;
                if (hexModel != null)
                {
                    return hexModel.Val;
                }

                // Get color from scheme
                A.SchemeColorValues runFontSchemeColor = aSolidFill.SchemeColor.Val.Value;
                return GetThemeColor(runFontSchemeColor);
            }

            // Get color from SHAPE level
            Shape parentShape = _portion.Paragraph.TextBox.AutoShape;
            if (parentShape.Placeholder is Placeholder placeholder)
            {
                P.Shape phShape = (P.Shape) placeholder.Shape.PShapeTreeChild;
                if (phShape.ShapeStyle != null)
                {
                    A.SchemeColorValues phShapeFontSchemeColor = phShape.ShapeStyle.FontReference.SchemeColor.Val.Value;
                    return GetThemeColor(phShapeFontSchemeColor);
                }

                if (placeholder.Type == PlaceholderType.Title)
                {
                    A.SchemeColorValues phTitleFontSchemeColor =
                        parentShape.SlideMaster.GetFontColorHexFromTitle(_paragraphLvl);
                    return GetThemeColor(phTitleFontSchemeColor);
                }

                if (placeholder.Type == PlaceholderType.Body)
                {
                    A.SchemeColorValues phBodyFontSchemeColor =
                        parentShape.SlideMaster.GetFontColorHexFromBody(_paragraphLvl);
                    return GetThemeColor(phBodyFontSchemeColor);
                }
            }

            P.Shape parentPShape = (P.Shape) parentShape.PShapeTreeChild;
            if (parentPShape.ShapeStyle != null)
            {
                A.SchemeColorValues shapeFontSchemeColor = parentPShape.ShapeStyle.FontReference.SchemeColor.Val.Value;
                return GetThemeColor(shapeFontSchemeColor);
            }

            A.SchemeColorValues bodyFontSchemeColor = parentShape.SlideMaster.GetFontColorHexFromBody(_paragraphLvl);
            return GetThemeColor(bodyFontSchemeColor);
        }

        private string GetThemeColor(A.SchemeColorValues fontSchemeColor)
        {
            A.ColorScheme themeAColorScheme =
                _portion.Paragraph.TextBox.AutoShape.ThemePart.Theme.ThemeElements.ColorScheme;
            return fontSchemeColor switch
            {
                A.SchemeColorValues.Dark1 => themeAColorScheme.Dark1Color.RgbColorModelHex != null
                    ? themeAColorScheme.Dark1Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Dark1Color.SystemColor.LastColor.Value,
                A.SchemeColorValues.Light1 => themeAColorScheme.Light1Color.RgbColorModelHex != null
                    ? themeAColorScheme.Light1Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Light1Color.SystemColor.LastColor.Value,
                A.SchemeColorValues.Dark2 => themeAColorScheme.Dark2Color.RgbColorModelHex != null
                    ? themeAColorScheme.Dark2Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Dark2Color.SystemColor.LastColor.Value,
                A.SchemeColorValues.Light2 => themeAColorScheme.Light2Color.RgbColorModelHex != null
                    ? themeAColorScheme.Light2Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Light2Color.SystemColor.LastColor.Value,
                A.SchemeColorValues.Accent1 => themeAColorScheme.Accent1Color.RgbColorModelHex != null
                    ? themeAColorScheme.Accent1Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Accent1Color.SystemColor.LastColor.Value,
                A.SchemeColorValues.Accent2 => themeAColorScheme.Accent2Color.RgbColorModelHex != null
                    ? themeAColorScheme.Accent2Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Accent2Color.SystemColor.LastColor.Value,
                A.SchemeColorValues.Accent3 => themeAColorScheme.Accent3Color.RgbColorModelHex != null
                    ? themeAColorScheme.Accent3Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Accent3Color.SystemColor.LastColor.Value,
                A.SchemeColorValues.Accent4 => themeAColorScheme.Accent4Color.RgbColorModelHex != null
                    ? themeAColorScheme.Accent4Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Accent4Color.SystemColor.LastColor.Value,
                A.SchemeColorValues.Accent5 => themeAColorScheme.Accent5Color.RgbColorModelHex != null
                    ? themeAColorScheme.Accent5Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Accent5Color.SystemColor.LastColor.Value,
                A.SchemeColorValues.Accent6 => themeAColorScheme.Accent6Color.RgbColorModelHex != null
                    ? themeAColorScheme.Accent6Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Accent6Color.SystemColor.LastColor.Value,
                A.SchemeColorValues.Hyperlink => themeAColorScheme.Hyperlink.RgbColorModelHex != null
                    ? themeAColorScheme.Hyperlink.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Hyperlink.SystemColor.LastColor.Value,
                _ => GetThemeMappedColor(fontSchemeColor)
            };

            string GetThemeMappedColor(A.SchemeColorValues fontSchemeColor)
            {
                P.ColorMap slideMasterPColorMap =
                    _portion.Paragraph.TextBox.AutoShape.SlideMaster.PSlideMaster.ColorMap;
                if (fontSchemeColor == A.SchemeColorValues.Text1)
                {
                    return GetThemeColorByString(slideMasterPColorMap.Text1.ToString());
                }

                if (fontSchemeColor == A.SchemeColorValues.Text2)
                {
                    return GetThemeColorByString(slideMasterPColorMap.Text2.ToString());
                }

                if (fontSchemeColor == A.SchemeColorValues.Background1)
                {
                    return GetThemeColorByString(slideMasterPColorMap.Background1.ToString());
                }

                return GetThemeColorByString(slideMasterPColorMap.Background2.ToString());
            }
        }

        private string GetThemeColorByString(string fontSchemeColor)
        {
            A.ColorScheme themeAColorScheme =
                _portion.Paragraph.TextBox.AutoShape.ThemePart.Theme.ThemeElements.ColorScheme;
            return fontSchemeColor switch
            {
                "dk1" => themeAColorScheme.Dark1Color.RgbColorModelHex != null
                    ? themeAColorScheme.Dark1Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Dark1Color.SystemColor.LastColor.Value,
                "lt1" => themeAColorScheme.Light1Color.RgbColorModelHex != null
                    ? themeAColorScheme.Light1Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Light1Color.SystemColor.LastColor.Value,
                "dk2" => themeAColorScheme.Dark2Color.RgbColorModelHex != null
                    ? themeAColorScheme.Dark2Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Dark2Color.SystemColor.LastColor.Value,
                "lt2" => themeAColorScheme.Light2Color.RgbColorModelHex != null
                    ? themeAColorScheme.Light2Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Light2Color.SystemColor.LastColor.Value,
                "accent1" => themeAColorScheme.Accent1Color.RgbColorModelHex != null
                    ? themeAColorScheme.Accent1Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Accent1Color.SystemColor.LastColor.Value,
                "accent2" => themeAColorScheme.Accent2Color.RgbColorModelHex != null
                    ? themeAColorScheme.Accent2Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Accent2Color.SystemColor.LastColor.Value,
                "accent3" => themeAColorScheme.Accent3Color.RgbColorModelHex != null
                    ? themeAColorScheme.Accent3Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Accent3Color.SystemColor.LastColor.Value,
                "accent4" => themeAColorScheme.Accent4Color.RgbColorModelHex != null
                    ? themeAColorScheme.Accent4Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Accent4Color.SystemColor.LastColor.Value,
                "accent5" => themeAColorScheme.Accent5Color.RgbColorModelHex != null
                    ? themeAColorScheme.Accent5Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Accent5Color.SystemColor.LastColor.Value,
                "accent6" => themeAColorScheme.Accent6Color.RgbColorModelHex != null
                    ? themeAColorScheme.Accent6Color.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Accent6Color.SystemColor.LastColor.Value,
                _ => themeAColorScheme.Hyperlink.RgbColorModelHex != null // hlink
                    ? themeAColorScheme.Hyperlink.RgbColorModelHex.Val.Value
                    : themeAColorScheme.Hyperlink.SystemColor.LastColor.Value
            };
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

        private void SetItalicFlag(bool value)
        {
            A.RunProperties aRunPr = _aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunPr != null)
            {
                aRunPr.Italic = new BooleanValue(value);
            }
            else
            {
                A.EndParagraphRunProperties aEndParaRPr = _aText.Parent.NextSibling<A.EndParagraphRunProperties>();
                if (aEndParaRPr != null)
                {
                    aEndParaRPr.Italic = new BooleanValue(value);
                }
                else
                {
                    aRunPr = new A.RunProperties {Italic = new BooleanValue(value)};
                    _aText.Parent.InsertAt(aRunPr, 0); // append to <a:r>
                }
            }
        }

        private string GetName()
        {
            const string majorLatinFont = "+mj-lt";
            if (_latinFont.Value.Typeface == majorLatinFont)
            {
                return _portion.Paragraph.TextBox.AutoShape.ThemePart.Theme.ThemeElements.FontScheme.MajorFont.LatinFont
                    .Typeface;
            }

            return _latinFont.Value.Typeface;
        }

        private void SetName(string fontName)
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

        private A.LatinFont GetALatinFont()
        {
            A.RunProperties aRunProperties = _aText.Parent.GetFirstChild<A.RunProperties>();
            A.LatinFont aLatinFont = aRunProperties?.GetFirstChild<A.LatinFont>();

            if (aLatinFont != null)
            {
                return aLatinFont;
            }

            if (TryGetFontDataFromPlaceholder(out FontData phFontData))
            {
                if (phFontData.ALatinFont != null)
                {
                    return phFontData.ALatinFont;
                }
            }

            // Get from theme
            return _portion.Paragraph.TextBox.AutoShape.ThemePart.Theme.ThemeElements.FontScheme.MinorFont.LatinFont;
        }

        private int GetSize()
        {
            Int32Value aRunPrFontSize = _portion.AText.Parent.GetFirstChild<A.RunProperties>()?.FontSize;
            if (aRunPrFontSize != null)
            {
                return aRunPrFontSize.Value;
            }

            Shape autoShape = _portion.Paragraph.TextBox.AutoShape;
            int paragraphLvl = _portion.Paragraph.Level;

            // Try get font size from placeholder
            if (autoShape.Placeholder != null)
            {
                Placeholder placeholder = (Placeholder) autoShape.Placeholder;
                IAutoShapeInternal placeholderAutoShape = (IAutoShapeInternal) placeholder.Shape;
                if (placeholderAutoShape != null &&
                    placeholderAutoShape.TryGetFontData(paragraphLvl, out FontData fontDataPlaceholder))
                {
                    if (fontDataPlaceholder.FontSize != null)
                    {
                        return fontDataPlaceholder.FontSize;
                    }
                }

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

        private bool GetBoldFlag()
        {
            A.RunProperties aRunProperties = _aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunProperties == null)
            {
                return false;
            }

            if (aRunProperties.Bold != null && aRunProperties.Bold == true)
            {
                return true;
            }

            if (TryGetFontDataFromPlaceholder(out FontData phFontData))
            {
                if (phFontData.IsBold != null)
                {
                    return phFontData.IsBold.Value;
                }
            }

            return false;
        }

        private bool TryGetFontDataFromPlaceholder(out FontData phFontData)
        {
            if (_portion.Paragraph.TextBox.AutoShape.Placeholder is Placeholder placeholder)
            {
                int paragraphLvl = _portion.Paragraph.Level;
                IAutoShapeInternal placeholderAutoShape = (IAutoShapeInternal) placeholder.Shape;
                if (placeholder.Shape != null && placeholderAutoShape.TryGetFontData(paragraphLvl, out phFontData))
                {
                    return true;
                }
            }

            phFontData = null;
            return false;
        }

        private bool GetItalicFlag()
        {
            A.RunProperties aRunProperties = _aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunProperties == null)
            {
                return false;
            }

            if (aRunProperties.Italic != null && aRunProperties.Italic == true)
            {
                return true;
            }

            if (TryGetFontDataFromPlaceholder(out FontData phFontData))
            {
                if (phFontData.IsItalic != null)
                {
                    return phFontData.IsItalic.Value;
                }
            }

            return false;
        }

        private void SetBoldFlag(bool value)
        {
            A.RunProperties aRunPr = _aText.Parent.GetFirstChild<A.RunProperties>();
            if (aRunPr != null)
            {
                aRunPr.Bold = new BooleanValue(value);
            }
            else
            {
                if (TryGetFontDataFromPlaceholder(out FontData phFontData))
                {
                    phFontData.IsBold = new BooleanValue(value);
                }
                else
                {
                    A.EndParagraphRunProperties aEndParaRPr = _aText.Parent.NextSibling<A.EndParagraphRunProperties>();
                    if (aEndParaRPr != null)
                    {
                        aEndParaRPr.Bold = new BooleanValue(value);
                    }
                    else
                    {
                        aRunPr = new A.RunProperties {Bold = new BooleanValue(value)};
                        _aText.Parent.InsertAt(aRunPr, 0); // append to <a:r>
                    }
                }
            }
        }

        #endregion Private Methods
    }
}
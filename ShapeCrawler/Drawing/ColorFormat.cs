using System;
using System.Drawing;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing
{
    internal class ColorFormat : IColorFormat
    {
        private readonly SCFont _font;
        private bool _initialized;
        private SCColorType _colorType;
        private Color _color;

        public ColorFormat(SCFont font)
        {
            _font = font;
        }

        public SCColorType ColorType => GetColorType();

        public Color Color
        {
            get => GetColor();
            set
            {

            }
        }

        private SCColorType GetColorType()
        {
            if (!_initialized)
            {
                InitColor();
            }

            return _colorType;
        }

        private Color GetColor()
        {
            if (!_initialized)
            {
                InitColor();
            }

            return _color;
        }

        private void InitColor()
        {
            _initialized = true;
            int paragraphLevel = _font.Portion.Paragraph.Level;

            // Try get color from PORTION level
            A.SolidFill aSolidFill = _font.Portion.AText.PreviousSibling<A.RunProperties>()?.SolidFill();
            if (aSolidFill != null)
            {
                // Try get solid color
                A.RgbColorModelHex hexModel = aSolidFill.RgbColorModelHex;
                if (hexModel != null)
                {
                    _colorType = SCColorType.RGB;
                    _color = ColorTranslator.FromHtml($"#{hexModel.Val}");
                    return;
                }

                // Try get from scheme color
                A.SchemeColor aSchemeColor = aSolidFill.SchemeColor;
                if (aSchemeColor != null)
                {
                    A.SchemeColorValues runFontSchemeColor = aSolidFill.SchemeColor.Val.Value;
                    string colorHex = GetThemeColor(runFontSchemeColor);
                    _colorType = SCColorType.Scheme;
                    _color = ColorTranslator.FromHtml($"#{colorHex}");
                    return;
                }

                A.PresetColor aPresetColor = aSolidFill.PresetColor;
                _colorType = SCColorType.Preset;
            }
            else
            {
                // Get color from SHAPE level
                Shape fontParentShape = _font.Portion.Paragraph.TextBox.AutoShape;
                string colorHexVariant;
                if (fontParentShape.Placeholder is Placeholder placeholder)
                {
                    FontData phFontData = new();
                    FontDataParser.GetFontDataFromPlaceholder(ref phFontData, _font.Portion.Paragraph);
                    if (phFontData.ARgbColorModelHex != null)
                    {
                        _colorType = SCColorType.RGB;
                        _color = ColorTranslator.FromHtml($"#{phFontData.ARgbColorModelHex.Val.Value}");
                        return;
                    }

                    if (phFontData.ASchemeColor != null)
                    {
                        colorHexVariant = GetThemeColor(phFontData.ASchemeColor.Val);
                        _colorType = SCColorType.Scheme;
                        _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                        return;
                    }

                    if (placeholder.Type == PlaceholderType.Title)
                    {
                        A.SchemeColorValues phTitleFontSchemeColor =
                            fontParentShape.SlideMaster.GetFontColorHexFromTitle(paragraphLevel);
                        colorHexVariant = GetThemeColor(phTitleFontSchemeColor);
                        _colorType = SCColorType.Scheme;
                        _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                        return;
                    }

                    if (placeholder.Type == PlaceholderType.Body)
                    {
                        A.SchemeColorValues phBodyFontSchemeColor = fontParentShape.SlideMaster.GetFontColorHexFromBody(paragraphLevel);
                        colorHexVariant = GetThemeColor(phBodyFontSchemeColor);
                        _colorType = SCColorType.Scheme;
                        _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                        return;
                    }
                }

                P.Shape parentPShape = (P.Shape)fontParentShape.PShapeTreeChild;
                if (parentPShape.ShapeStyle != null)
                {
                    A.SchemeColorValues shapeFontSchemeColor = parentPShape.ShapeStyle.FontReference.SchemeColor.Val.Value;
                    colorHexVariant = GetThemeColor(shapeFontSchemeColor);
                    _colorType = SCColorType.Scheme;
                    _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                    return;
                }

                FontData masterBodyFontData = fontParentShape.SlideMaster.BodyParaLvlToFontData[paragraphLevel];
                if (masterBodyFontData.ARgbColorModelHex != null)
                {
                    _colorType = SCColorType.RGB;
                    _color = ColorTranslator.FromHtml($"#{masterBodyFontData.ARgbColorModelHex.Val.Value}");
                    return;
                }

                string colorHex = GetThemeColor(masterBodyFontData.ASchemeColor.Val);
                _colorType = SCColorType.Scheme;
                _color = ColorTranslator.FromHtml($"#{colorHex}");
            }
        }

        private string GetThemeColor(A.SchemeColorValues fontSchemeColor)
        {
            A.ColorScheme themeAColorScheme =
                _font.Portion.Paragraph.TextBox.AutoShape.ThemePart.Theme.ThemeElements.ColorScheme;
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
                    _font.Portion.Paragraph.TextBox.AutoShape.SlideMaster.PSlideMaster.ColorMap;
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
                _font.Portion.Paragraph.TextBox.AutoShape.ThemePart.Theme.ThemeElements.ColorScheme;
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
    }
}
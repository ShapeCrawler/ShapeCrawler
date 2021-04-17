using System.Collections.Generic;
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
        private Color _color;
        private SCColorType _colorType;
        private bool _initialized;
        private static readonly Dictionary<A.PresetColorValues, Color> PresentColorToColor = new()
        {
            { A.PresetColorValues.AliceBlue, Color.AliceBlue },
            { A.PresetColorValues.AntiqueWhite, Color.AntiqueWhite },
            { A.PresetColorValues.Aqua, Color.Aqua },
            { A.PresetColorValues.Aquamarine, Color.Aquamarine },
            { A.PresetColorValues.Azure, Color.Azure },
            { A.PresetColorValues.Beige, Color.Beige },
            { A.PresetColorValues.Bisque, Color.Bisque },
            { A.PresetColorValues.Black, Color.Black },
            { A.PresetColorValues.BlanchedAlmond, Color.BlanchedAlmond },
            { A.PresetColorValues.Blue, Color.Blue },
            { A.PresetColorValues.BlueViolet, Color.BlueViolet },
            { A.PresetColorValues.Brown, Color.Brown },
            { A.PresetColorValues.BurlyWood, Color.BurlyWood },
            { A.PresetColorValues.CadetBlue, Color.CadetBlue },
            { A.PresetColorValues.Chartreuse, Color.Chartreuse },
            { A.PresetColorValues.Chocolate, Color.Chocolate },
            { A.PresetColorValues.Coral, Color.Coral },
            { A.PresetColorValues.CornflowerBlue, Color.CornflowerBlue },
            { A.PresetColorValues.Cornsilk, Color.Cornsilk },
            { A.PresetColorValues.Crimson, Color.Crimson },
            { A.PresetColorValues.Cyan, Color.Cyan },
            { A.PresetColorValues.DarkBlue, Color.DarkBlue },
            { A.PresetColorValues.DarkCyan, Color.DarkCyan },
            { A.PresetColorValues.DarkGoldenrod, Color.DarkGoldenrod },
            { A.PresetColorValues.DarkGray, Color.DarkGray },
            { A.PresetColorValues.DarkGreen, Color.DarkGreen },
            { A.PresetColorValues.DarkKhaki, Color.DarkKhaki },
            { A.PresetColorValues.DarkMagenta, Color.DarkMagenta },
            { A.PresetColorValues.DarkOliveGreen, Color.DarkOliveGreen },
            { A.PresetColorValues.DarkOrange, Color.DarkOrange },
            { A.PresetColorValues.DarkOrchid, Color.DarkOrchid },
            { A.PresetColorValues.DarkRed, Color.DarkRed },
            { A.PresetColorValues.DarkSalmon, Color.DarkSalmon },
            { A.PresetColorValues.DarkSeaGreen, Color.DarkSeaGreen },
            { A.PresetColorValues.DarkSlateBlue, Color.DarkSlateBlue },
            { A.PresetColorValues.DarkSlateGray, Color.DarkSlateGray },
            { A.PresetColorValues.DarkTurquoise, Color.DarkTurquoise },
            { A.PresetColorValues.DarkViolet, Color.DarkViolet },
            { A.PresetColorValues.DeepPink, Color.DeepPink },
            { A.PresetColorValues.DeepSkyBlue, Color.DeepSkyBlue },
            { A.PresetColorValues.DimGray, Color.DimGray },
            { A.PresetColorValues.DodgerBlue, Color.DodgerBlue },
            { A.PresetColorValues.Firebrick, Color.Firebrick },
            { A.PresetColorValues.FloralWhite, Color.FloralWhite },
            { A.PresetColorValues.ForestGreen, Color.ForestGreen },
            { A.PresetColorValues.Fuchsia, Color.Fuchsia },
            { A.PresetColorValues.Gainsboro, Color.Gainsboro },
            { A.PresetColorValues.GhostWhite, Color.GhostWhite },
            { A.PresetColorValues.Gold, Color.Gold },
            { A.PresetColorValues.Goldenrod, Color.Goldenrod },
            { A.PresetColorValues.Gray, Color.Gray },
            { A.PresetColorValues.Green, Color.Green },
            { A.PresetColorValues.GreenYellow, Color.GreenYellow },
            { A.PresetColorValues.Honeydew, Color.Honeydew },
            { A.PresetColorValues.HotPink, Color.HotPink },
            { A.PresetColorValues.IndianRed, Color.IndianRed },
            { A.PresetColorValues.Indigo, Color.Indigo },
            { A.PresetColorValues.Ivory, Color.Ivory },
            { A.PresetColorValues.Khaki, Color.Khaki },
            { A.PresetColorValues.Lavender, Color.Lavender },
            { A.PresetColorValues.LavenderBlush, Color.LavenderBlush },
            { A.PresetColorValues.LawnGreen, Color.LawnGreen },
            { A.PresetColorValues.LemonChiffon, Color.LemonChiffon },
            { A.PresetColorValues.LightBlue, Color.LightBlue },
            { A.PresetColorValues.LightCoral, Color.LightCoral },
            { A.PresetColorValues.LightCyan, Color.LightCyan },
            { A.PresetColorValues.LightGoldenrodYellow, Color.LightGoldenrodYellow },
            { A.PresetColorValues.LightGray, Color.LightGray },
            { A.PresetColorValues.LightGreen, Color.LightGreen },
            { A.PresetColorValues.LightPink, Color.LightPink },
            { A.PresetColorValues.LightSalmon, Color.LightSalmon },
            { A.PresetColorValues.LightSeaGreen, Color.LightSeaGreen },
            { A.PresetColorValues.LightSkyBlue, Color.LightSkyBlue },
            { A.PresetColorValues.LightSlateGray, Color.LightSlateGray },
            { A.PresetColorValues.LightSteelBlue, Color.LightSteelBlue },
            { A.PresetColorValues.LightYellow, Color.LightYellow },
            { A.PresetColorValues.Lime, Color.Lime },
            { A.PresetColorValues.LimeGreen, Color.LimeGreen },
            { A.PresetColorValues.Linen, Color.Linen },
            { A.PresetColorValues.Magenta, Color.Magenta },
            { A.PresetColorValues.Maroon, Color.Maroon },
            { A.PresetColorValues.MediumBlue, Color.MediumBlue },
            { A.PresetColorValues.MediumOrchid, Color.MediumOrchid },
            { A.PresetColorValues.MediumPurple, Color.MediumPurple },
            { A.PresetColorValues.MediumSeaGreen, Color.MediumSeaGreen },
            { A.PresetColorValues.MediumSlateBlue, Color.MediumSlateBlue },
            { A.PresetColorValues.MediumSpringGreen, Color.MediumSpringGreen },
            { A.PresetColorValues.MediumTurquoise, Color.MediumTurquoise },
            { A.PresetColorValues.MediumVioletRed, Color.MediumVioletRed },
            { A.PresetColorValues.MidnightBlue, Color.MidnightBlue },
            { A.PresetColorValues.MintCream, Color.MintCream },
            { A.PresetColorValues.MistyRose, Color.MistyRose },
            { A.PresetColorValues.Moccasin, Color.Moccasin },
            { A.PresetColorValues.NavajoWhite, Color.NavajoWhite },
            { A.PresetColorValues.Navy, Color.Navy },
            { A.PresetColorValues.OldLace, Color.OldLace },
            { A.PresetColorValues.Olive, Color.Olive },
            { A.PresetColorValues.OliveDrab, Color.OliveDrab },
            { A.PresetColorValues.Orange, Color.Orange },
            { A.PresetColorValues.OrangeRed, Color.OrangeRed },
            { A.PresetColorValues.Orchid, Color.Orchid },
            { A.PresetColorValues.PaleGoldenrod, Color.PaleGoldenrod },
            { A.PresetColorValues.PaleGreen, Color.PaleGreen },
            { A.PresetColorValues.PaleTurquoise, Color.PaleTurquoise },
            { A.PresetColorValues.PaleVioletRed, Color.PaleVioletRed },
            { A.PresetColorValues.PapayaWhip, Color.PapayaWhip },
            { A.PresetColorValues.PeachPuff, Color.PeachPuff },
            { A.PresetColorValues.Peru, Color.Peru },
            { A.PresetColorValues.Pink, Color.Pink },
            { A.PresetColorValues.Plum, Color.Plum },
            { A.PresetColorValues.PowderBlue, Color.PowderBlue },
            { A.PresetColorValues.Purple, Color.Purple },
            { A.PresetColorValues.Red, Color.Red },
            { A.PresetColorValues.RosyBrown, Color.RosyBrown },
            { A.PresetColorValues.RoyalBlue, Color.RoyalBlue },
            { A.PresetColorValues.SaddleBrown, Color.SaddleBrown },
            { A.PresetColorValues.Salmon, Color.Salmon },
            { A.PresetColorValues.SandyBrown, Color.SandyBrown },
            { A.PresetColorValues.SeaGreen, Color.SeaGreen },
            { A.PresetColorValues.SeaShell, Color.SeaShell },
            { A.PresetColorValues.Sienna, Color.Sienna },
            { A.PresetColorValues.Silver, Color.Silver },
            { A.PresetColorValues.SkyBlue, Color.SkyBlue },
            { A.PresetColorValues.SlateBlue, Color.SlateBlue },
            { A.PresetColorValues.SlateGray, Color.SlateGray },
            { A.PresetColorValues.Snow, Color.Snow },
            { A.PresetColorValues.SpringGreen, Color.SpringGreen },
            { A.PresetColorValues.SteelBlue, Color.SteelBlue },
            { A.PresetColorValues.Tan, Color.Tan },
            { A.PresetColorValues.Teal, Color.Teal },
            { A.PresetColorValues.Thistle, Color.Thistle },
            { A.PresetColorValues.Tomato, Color.Tomato },
            { A.PresetColorValues.Turquoise, Color.Turquoise },
            { A.PresetColorValues.Violet, Color.Violet },
            { A.PresetColorValues.Wheat, Color.Wheat },
            { A.PresetColorValues.White, Color.White },
            { A.PresetColorValues.WhiteSmoke, Color.WhiteSmoke },
            { A.PresetColorValues.Yellow, Color.Yellow },
            { A.PresetColorValues.YellowGreen, Color.YellowGreen }
        };

        public ColorFormat(SCFont font)
        {
            _font = font;
        }

        public SCColorType ColorType => GetColorType();

        public Color Color
        {
            get => GetColor();
            set { }
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
            string colorHexVariant;

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

                // Try get scheme color
                A.SchemeColor aSchemeColor = aSolidFill.SchemeColor;
                if (aSchemeColor != null)
                {
                    colorHexVariant = GetHexVariantByScheme(aSchemeColor.Val);
                    _colorType = SCColorType.Scheme;
                    _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                    return;
                }

                // Try get system color
                A.SystemColor aSystemColor = aSolidFill.SystemColor;
                if (aSystemColor != null)
                {
                    colorHexVariant = aSystemColor.LastColor;
                    _colorType = SCColorType.System;
                    _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                    return;
                }

                A.PresetColor aPresetColor = aSolidFill.PresetColor;
                _colorType = SCColorType.Preset;
                _color = PresentColorToColor[aPresetColor.Val.Value];
            }
            else
            {
                // Get color from SHAPE level
                Shape fontParentShape = _font.Portion.Paragraph.TextBox.AutoShape;
                if (fontParentShape.Placeholder is Placeholder placeholder)
                {
                    FontData phFontData = new();
                    FontDataParser.GetFontDataFromPlaceholder(ref phFontData, _font.Portion.Paragraph);
                    if (phFontData.ARgbColorModelHex != null)
                    {
                        colorHexVariant = phFontData.ARgbColorModelHex.Val;
                        _colorType = SCColorType.RGB;
                        _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                        return;
                    }

                    if (phFontData.ASchemeColor != null)
                    {
                        colorHexVariant = GetHexVariantByScheme(phFontData.ASchemeColor.Val);
                        _colorType = SCColorType.Scheme;
                        _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                        return;
                    }

                    if (placeholder.Type == PlaceholderType.Title)
                    {
                        FontData masterTitleFontData = fontParentShape.SlideMaster.TitleParaLvlToFontData[paragraphLevel];
                        if (masterTitleFontData.ASchemeColor != null)
                        {
                            colorHexVariant = GetHexVariantByScheme(masterTitleFontData.ASchemeColor.Val);
                            _colorType = SCColorType.Scheme;
                            _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                        }
                        else
                        {
                            colorHexVariant = masterTitleFontData.ARgbColorModelHex.Val;
                            _colorType = SCColorType.RGB;
                            _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                        }

                        return;
                    }

                    if (placeholder.Type == PlaceholderType.Body)
                    {
                        A.SchemeColorValues phBodyFontSchemeColor =
                            fontParentShape.SlideMaster.GetFontColorHexFromBody(paragraphLevel);
                        colorHexVariant = GetHexVariantByScheme(phBodyFontSchemeColor);
                        _colorType = SCColorType.Scheme;
                        _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                        return;
                    }
                }

                P.Shape parentPShape = (P.Shape) fontParentShape.PShapeTreeChild;
                if (parentPShape.ShapeStyle != null)
                {
                    A.SchemeColorValues shapeFontSchemeColor =
                        parentPShape.ShapeStyle.FontReference.SchemeColor.Val.Value;
                    colorHexVariant = GetHexVariantByScheme(shapeFontSchemeColor);
                    _colorType = SCColorType.Scheme;
                    _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                    return;
                }

                // Try get from Slide Master
                
                FontData masterBodyFontData = fontParentShape.SlideMaster.BodyParaLvlToFontData[paragraphLevel];
                if (masterBodyFontData.ARgbColorModelHex != null)
                {
                    _colorType = SCColorType.RGB;
                    _color = ColorTranslator.FromHtml($"#{masterBodyFontData.ARgbColorModelHex.Val.Value}");
                    return;
                }

                if (masterBodyFontData.ASchemeColor != null)
                {
                    colorHexVariant = GetHexVariantByScheme(masterBodyFontData.ASchemeColor.Val);
                    _colorType = SCColorType.Scheme;
                    _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                    return;
                }

                // Try get from presentation global
                if (fontParentShape.Presentation.ParaLvlToFontData.TryGetValue(paragraphLevel, out FontData preFontData))
                {
                    FontData preDefaultTxtStyleFontData = fontParentShape.Presentation.ParaLvlToFontData[paragraphLevel];
                    colorHexVariant = GetHexVariantByScheme(preDefaultTxtStyleFontData.ASchemeColor.Val);
                    _colorType = SCColorType.Scheme;
                    _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                    return;
                }

                // Get default
                colorHexVariant = GetThemeMappedColor(A.SchemeColorValues.Text1);
                _colorType = SCColorType.Scheme;
                _color = ColorTranslator.FromHtml($"#{colorHexVariant}");
            }
        }

        private string GetHexVariantByScheme(A.SchemeColorValues fontSchemeColor)
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
        }

        private string GetThemeMappedColor(A.SchemeColorValues fontSchemeColor)
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
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
        private readonly SCFont font;
        private readonly Shape parentShape;
        private Color color;
        private SCColorType colorType;
        private bool initialized;

        internal ColorFormat(SCFont font)
        {
            this.font = font;
            this.parentShape = (Shape)font.Portion.ParentParagraph.ParentTextBox.ParentTextBoxContainer;
        }

        public SCColorType ColorType => this.GetColorType();

        public Color Color
        {
            get => this.GetColor();
            set { }
        }

        private SCColorType GetColorType()
        {
            if (!this.initialized)
            {
                this.InitColor();
            }

            return this.colorType;
        }

        private Color GetColor()
        {
            if (!this.initialized)
            {
                this.InitColor();
            }

            return this.color;
        }

        private void InitColor()
        {
            this.initialized = true;
            int paragraphLevel = this.font.Portion.ParentParagraph.Level;
            string colorHexVariant;

            // Try get color from PORTION level
            A.SolidFill aSolidFill = this.font.Portion.AText.Parent.GetFirstChild<A.RunProperties>()?.SolidFill();
            if (aSolidFill != null)
            {
                A.RgbColorModelHex hexModel = aSolidFill.RgbColorModelHex;
                if (hexModel != null)
                {
                    colorHexVariant = hexModel.Val;
                    this.colorType = SCColorType.RGB;
                    this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                    return;
                }

                A.SchemeColor aSchemeColor = aSolidFill.SchemeColor;
                if (aSchemeColor != null)
                {
                    colorHexVariant = this.GetHexVariantByScheme(aSchemeColor.Val);
                    this.colorType = SCColorType.Scheme;
                    this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                    return;
                }

                A.SystemColor aSystemColor = aSolidFill.SystemColor;
                if (aSystemColor != null)
                {
                    colorHexVariant = aSystemColor.LastColor;
                    this.colorType = SCColorType.System;
                    this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                    return;
                }

                A.PresetColor aPresetColor = aSolidFill.PresetColor;
                this.colorType = SCColorType.Preset;
                this.color = Color.FromName(aPresetColor.Val.Value.ToString());
            }
            else
            {
                // Get color from SHAPE level
                FontData masterBodyFontData;
                if (this.parentShape.Placeholder is Placeholder placeholder)
                {
                    FontData placeholderFontData = new ();
                    FontDataParser.GetFontDataFromPlaceholder(ref placeholderFontData, this.font.Portion.ParentParagraph);
                    if (placeholderFontData.ARgbColorModelHex != null)
                    {
                        colorHexVariant = placeholderFontData.ARgbColorModelHex.Val;
                        this.colorType = SCColorType.RGB;
                        this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                        return;
                    }

                    if (placeholderFontData.ASchemeColor != null)
                    {
                        colorHexVariant = this.GetHexVariantByScheme(placeholderFontData.ASchemeColor.Val);
                        this.colorType = SCColorType.Scheme;
                        this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                        return;
                    }

                    if (placeholderFontData.ASystemColor != null)
                    {
                        colorHexVariant = placeholderFontData.ASystemColor.LastColor;
                        this.colorType = SCColorType.System;
                        this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                        return;
                    }

                    if (placeholderFontData.APresetColor != null)
                    {
                        this.colorType = SCColorType.Preset;
                        this.color = Color.FromName(placeholderFontData.APresetColor.Val.Value.ToString());
                        return;
                    }

                    switch (placeholder.Type)
                    {
                        case PlaceholderType.Title:
                        {
                            FontData masterTitleFontData =
                                this.parentShape.SlideMaster.TitleParaLvlToFontData.ContainsKey(paragraphLevel)
                                    ? this.parentShape.SlideMaster.TitleParaLvlToFontData[paragraphLevel]
                                    : this.parentShape.SlideMaster.TitleParaLvlToFontData[1];
                            if (masterTitleFontData.ASchemeColor != null)
                            {
                                colorHexVariant = this.GetHexVariantByScheme(masterTitleFontData.ASchemeColor.Val);
                                this.colorType = SCColorType.Scheme;
                                this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                            }
                            else if (masterTitleFontData.ARgbColorModelHex != null)
                            {
                                colorHexVariant = masterTitleFontData.ARgbColorModelHex.Val;
                                this.colorType = SCColorType.RGB;
                                this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                            }
                            else
                            {
                                // Get default
                                colorHexVariant = this.GetThemeMappedColor(A.SchemeColorValues.Text1);
                                this.colorType = SCColorType.Scheme;
                                this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                            }

                            return;
                        }

                        case PlaceholderType.Body:
                        {
                            masterBodyFontData = this.parentShape.SlideMaster.BodyParaLvlToFontData[paragraphLevel];
                            if (masterBodyFontData.ASchemeColor != null)
                            {
                                A.SchemeColorValues phBodyFontSchemeColor = masterBodyFontData.ASchemeColor.Val;
                                colorHexVariant = this.GetHexVariantByScheme(phBodyFontSchemeColor);
                                this.colorType = SCColorType.Scheme;
                                this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                            }
                            else
                            {
                                colorHexVariant = masterBodyFontData.ARgbColorModelHex.Val;
                                this.colorType = SCColorType.RGB;
                                this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                            }

                            return;
                        }
                    }
                }

                // Shape level
                P.Shape parentPShape = (P.Shape)this.parentShape.PShapeTreeChild;
                if (parentPShape.ShapeStyle != null)
                {
                    A.RgbColorModelHex hexModel = parentPShape.ShapeStyle.FontReference.RgbColorModelHex;
                    if (hexModel != null)
                    {
                        colorHexVariant = hexModel.Val;
                        this.colorType = SCColorType.RGB;
                        this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                        return;
                    }

                    A.SchemeColor aSchemeColor = parentPShape.ShapeStyle.FontReference.SchemeColor;
                    if (aSchemeColor != null)
                    {
                        colorHexVariant = this.GetHexVariantByScheme(aSchemeColor.Val);
                        this.colorType = SCColorType.Scheme;
                        this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                        return;
                    }

                    A.SystemColor aSystemColor = parentPShape.ShapeStyle.FontReference.SystemColor;
                    if (aSystemColor != null)
                    {
                        colorHexVariant = aSystemColor.LastColor;
                        this.colorType = SCColorType.System;
                        this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                        return;
                    }

                    A.PresetColor aPresetColor = parentPShape.ShapeStyle.FontReference.PresetColor;
                    this.colorType = SCColorType.Preset;
                    this.color = Color.FromName(aPresetColor.Val.Value.ToString());
                }

                // Slide Master level
                masterBodyFontData = this.parentShape.SlideMaster.BodyParaLvlToFontData[paragraphLevel];
                if (masterBodyFontData.ARgbColorModelHex != null)
                {
                    colorHexVariant = masterBodyFontData.ARgbColorModelHex.Val.Value;
                    this.colorType = SCColorType.RGB;
                    this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                    return;
                }

                if (masterBodyFontData.ASchemeColor != null)
                {
                    colorHexVariant = this.GetHexVariantByScheme(masterBodyFontData.ASchemeColor.Val);
                    this.colorType = SCColorType.Scheme;
                    this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                    return;
                }

                // Presentation level
                if (this.parentShape.ParentPresentation.ParaLvlToFontData.TryGetValue(
                    paragraphLevel,
                    out FontData preFontData))
                {
                    colorHexVariant = this.GetHexVariantByScheme(preFontData.ASchemeColor.Val);
                    this.colorType = SCColorType.Scheme;
                    this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                    return;
                }

                // Get default
                colorHexVariant = this.GetThemeMappedColor(A.SchemeColorValues.Text1);
                this.colorType = SCColorType.Scheme;
                this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
            }
        }

        private string GetHexVariantByScheme(A.SchemeColorValues fontSchemeColor)
        {
            A.ColorScheme themeAColorScheme = this.parentShape.ThemePart.Theme.ThemeElements.ColorScheme;
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
                _ => this.GetThemeMappedColor(fontSchemeColor)
            };
        }

        private string GetThemeMappedColor(A.SchemeColorValues fontSchemeColor)
        {
            P.ColorMap slideMasterPColorMap = this.parentShape.SlideMaster.PSlideMaster.ColorMap;
            if (fontSchemeColor == A.SchemeColorValues.Text1)
            {
                return this.GetThemeColorByString(slideMasterPColorMap.Text1.ToString());
            }

            if (fontSchemeColor == A.SchemeColorValues.Text2)
            {
                return this.GetThemeColorByString(slideMasterPColorMap.Text2.ToString());
            }

            if (fontSchemeColor == A.SchemeColorValues.Background1)
            {
                return this.GetThemeColorByString(slideMasterPColorMap.Background1.ToString());
            }

            return this.GetThemeColorByString(slideMasterPColorMap.Background2.ToString());
        }

        private string GetThemeColorByString(string fontSchemeColor)
        {
            A.ColorScheme themeAColorScheme = this.parentShape.ThemePart.Theme.ThemeElements.ColorScheme;
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
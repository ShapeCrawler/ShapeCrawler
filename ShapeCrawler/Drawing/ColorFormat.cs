﻿using System.Collections.Generic;
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
        private readonly SCFont parentFont;
        private readonly ITextBoxContainer parentTextBoxContainer;
        private readonly SCSlideMaster parentSlideMaster;
        private Color color;
        private SCColorType colorType;
        private bool initialized;

        internal ColorFormat(SCFont parentFont)
        {
            this.parentFont = parentFont;
            this.parentTextBoxContainer = parentFont.ParentPortion.ParentParagraph.ParentTextBox.ParentTextBoxContainer;
            this.parentSlideMaster = this.parentTextBoxContainer.ParentSlideMaster;
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
            SCPortion parentPortion = this.parentFont.ParentPortion;
            SCParagraph parentParagraph = parentPortion.ParentParagraph;
            int paragraphLevel = parentParagraph.Level;

            A.SolidFill aSolidFill = parentPortion.AText.Parent.GetFirstChild<A.RunProperties>()?.SolidFill();
            if (aSolidFill != null)
            {
                this.FromRunSolidFill(aSolidFill);
            }
            else
            {
                if (this.TryFromTextBody(parentParagraph))
                {
                    return;
                }

                if (this.TryFromShapeFontReference())
                {
                    return;
                }

                if (this.TryFromPlaceholder(paragraphLevel))
                {
                    return;
                }

                FontData masterBodyFontData = this.parentSlideMaster.BodyParaLvlToFontData[paragraphLevel];
                if (this.TryFromFontData(masterBodyFontData))
                {
                    return;
                }

                // Presentation level
                string colorHexVariant;
                if (this.parentSlideMaster.ParentPresentation.ParaLvlToFontData.TryGetValue(paragraphLevel, out FontData preFontData))
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

        private bool TryFromTextBody(SCParagraph paragraph)
        {
            A.ListStyle txBodyListStyle = paragraph.ParentTextBox.APTextBody.GetFirstChild<A.ListStyle>();
            Dictionary<int, FontData> paraLvlToFontData = FontDataParser.FromCompositeElement(txBodyListStyle);
            if (!paraLvlToFontData.TryGetValue(paragraph.Level, out FontData txBodyFontData))
            {
                return false;
            }

            return this.TryFromFontData(txBodyFontData);
        }

        private bool TryFromShapeFontReference()
        {
            if (this.parentTextBoxContainer is Shape parentShape)
            {
                P.Shape parentPShape = (P.Shape) parentShape.SdkPShapeTreeChild;
                if (parentPShape.ShapeStyle == null)
                {
                    return false;
                }

                A.FontReference aFontReference = parentPShape.ShapeStyle.FontReference;
                FontData fontReferenceFontData = new()
                {
                    ARgbColorModelHex = aFontReference.RgbColorModelHex,
                    ASchemeColor = aFontReference.SchemeColor,
                    ASystemColor = aFontReference.SystemColor,
                    APresetColor = aFontReference.PresetColor
                };

                return this.TryFromFontData(fontReferenceFontData);
            }

            return false;
        }

        private bool TryFromPlaceholder(int paragraphLevel)
        {
            if (this.parentTextBoxContainer.Placeholder is not Placeholder placeholder)
            {
                return false;
            }

            FontData placeholderFontData = new ();
            FontDataParser.GetFontDataFromPlaceholder(ref placeholderFontData, this.parentFont.ParentPortion.ParentParagraph);
            if (this.TryFromFontData(placeholderFontData))
            {
                return true;
            }

            switch (placeholder.Type)
            {
                case PlaceholderType.Title:
                {
                    Dictionary<int, FontData> titleParaLvlToFontData = this.parentSlideMaster.TitleParaLvlToFontData;
                    FontData masterTitleFontData = titleParaLvlToFontData.ContainsKey(paragraphLevel)
                        ? titleParaLvlToFontData[paragraphLevel]
                        : titleParaLvlToFontData[1];
                    if (this.TryFromFontData(masterTitleFontData))
                    {
                        return true;
                    }

                    break;
                }

                case PlaceholderType.Body:
                {
                    Dictionary<int, FontData> bodyParaLvlToFontData = this.parentSlideMaster.BodyParaLvlToFontData;
                    FontData masterBodyFontData = bodyParaLvlToFontData[paragraphLevel];
                    if (this.TryFromFontData(masterBodyFontData))
                    {
                        return true;
                    }

                    break;
                }
            }

            return false;
        }

        private bool TryFromFontData(FontData fontData)
        {
            string colorHexVariant;
            if (fontData.ARgbColorModelHex != null)
            {
                colorHexVariant = fontData.ARgbColorModelHex.Val;
                this.colorType = SCColorType.RGB;
                this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                return true;
            }

            if (fontData.ASchemeColor != null)
            {
                colorHexVariant = this.GetHexVariantByScheme(fontData.ASchemeColor.Val);
                this.colorType = SCColorType.Scheme;
                this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                return true;
            }

            if (fontData.ASystemColor != null)
            {
                colorHexVariant = fontData.ASystemColor.LastColor;
                this.colorType = SCColorType.System;
                this.color = ColorTranslator.FromHtml($"#{colorHexVariant}");
                return true;
            }

            if (fontData.APresetColor != null)
            {
                this.colorType = SCColorType.Preset;
                this.color = Color.FromName(fontData.APresetColor.Val.Value.ToString());
                return true;
            }

            return false;
        }

        private void FromRunSolidFill(A.SolidFill aSolidFill)
        {
            A.RgbColorModelHex hexModel = aSolidFill.RgbColorModelHex;
            string colorHexVariant;
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

        private string GetHexVariantByScheme(A.SchemeColorValues fontSchemeColor)
        {
            A.ColorScheme themeAColorScheme = this.parentSlideMaster.ThemePart.Theme.ThemeElements.ColorScheme;
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
            P.ColorMap slideMasterPColorMap = this.parentSlideMaster.PSlideMaster.ColorMap;
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
            A.ColorScheme themeAColorScheme = this.parentSlideMaster.ThemePart.Theme.ThemeElements.ColorScheme;
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
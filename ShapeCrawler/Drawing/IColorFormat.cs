using System.Collections.Generic;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Services;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a color.
/// </summary>
public interface IColorFormat
{
    /// <summary>
    ///     Gets color type.
    /// </summary>
    SCColorType ColorType { get; }

    /// <summary>
    ///     Gets color hexadecimal representation.
    /// </summary>
    string ColorHex { get; } // TODO: combine with SetColorByHex?

    /// <summary>
    ///     Sets solid color by its hexadecimal representation.
    /// </summary>
    void SetColorByHex(string hex);
}

internal sealed class ColorFormat : IColorFormat
{
    private readonly SCFont font;
    private readonly ITextFrameContainer textFrameContainer;
    private readonly SCSlideMaster parentSlideMaster;
    private bool initialized;
    private string? hexColor;
    private SCColorType colorType;

    internal ColorFormat(SCFont font)
    {
        this.font = font;
        this.textFrameContainer = font.ParentPortion.ParentParagraph.ParentTextFrame.TextFrameContainer;
        var shape = this.textFrameContainer.Shape;
        this.parentSlideMaster = shape.SlideMasterInternal;
    }

    public SCColorType ColorType => this.GetColorType();

    public string ColorHex => this.GetColorHex();

    public void SetColorByHex(string hex)
    {
        var portion = this.font.ParentPortion;
        var aTextContainer = portion.AText.Parent!;
        var aRunProperties = aTextContainer.GetFirstChild<A.RunProperties>() ?? aTextContainer.AddRunProperties();

        var aSolidFill = aRunProperties.GetASolidFill();
        aSolidFill?.Remove();

        hex = hex.Substring(1); // to skip '#'
        var rgbColorModelHex = new A.RgbColorModelHex { Val = hex };
        aSolidFill = new A.SolidFill();
        aSolidFill.Append(rgbColorModelHex);
        aRunProperties.Append(aSolidFill);
    }

    private SCColorType GetColorType()
    {
        if (!this.initialized)
        {
            this.InitializeColor();
        }

        return this.colorType;
    }

    private string GetColorHex()
    {
        if (!this.initialized)
        {
            this.InitializeColor();
        }

        return this.hexColor!;
    }

    private void InitializeColor()
    {
        this.initialized = true;
        var portion = this.font.ParentPortion;
        var aSolidFill = portion.AText.Parent!.GetFirstChild<A.RunProperties>()?.GetASolidFill();
        if (aSolidFill != null)
        {
            var typeAndHex = HexParser.FromSolidFill(aSolidFill, this.parentSlideMaster);
            this.colorType = typeAndHex.Item1;
            this.hexColor = typeAndHex.Item2;
        }
        else
        {
            var paragraph = portion.ParentParagraph;
            var paragraphLevel = paragraph.Level;
            if (this.TryFromTextBody(paragraph))
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
            if (this.parentSlideMaster.PresentationInternal.ParaLvlToFontData.TryGetValue(paragraphLevel, out var preFontData))
            {
                colorHexVariant = this.GetHexVariantByScheme(preFontData.ASchemeColor!.Val!);
                this.colorType = SCColorType.Scheme;
                this.hexColor = colorHexVariant;
                return;
            }

            // Get default
            colorHexVariant = this.GetThemeMappedColor(A.SchemeColorValues.Text1);
            this.colorType = SCColorType.Scheme;
            this.hexColor = colorHexVariant;
        }
    }

    private bool TryFromTextBody(SCParagraph paragraph)
    {
        var txBodyListStyle = paragraph.ParentTextFrame.TextBodyElement!.GetFirstChild<A.ListStyle>();
        var paraLvlToFontData = FontDataParser.FromCompositeElement(txBodyListStyle!);
        if (!paraLvlToFontData.TryGetValue(paragraph.Level, out var txBodyFontData))
        {
            return false;
        }

        return this.TryFromFontData(txBodyFontData);
    }

    private bool TryFromShapeFontReference()
    {
        if (this.textFrameContainer is Shape parentShape)
        {
            var parentPShape = (P.Shape)parentShape.PShapeTreesChild;
            if (parentPShape.ShapeStyle == null)
            {
                return false;
            }

            var aFontReference = parentPShape.ShapeStyle.FontReference!;
            var fontReferenceFontData = new FontData()
            {
                ARgbColorModelHex = aFontReference.RgbColorModelHex,
                ASchemeColor = aFontReference.SchemeColor,
                APresetColor = aFontReference.PresetColor
            };

            return this.TryFromFontData(fontReferenceFontData);
        }

        return false;
    }

    private bool TryFromPlaceholder(int paragraphLevel)
    {
        if (this.textFrameContainer.Shape.Placeholder is not Placeholder placeholder)
        {
            return false;
        }

        var phFontData = new FontData();
        var paragraph = this.font.ParentPortion.ParentParagraph;
        FontDataParser.GetFontDataFromPlaceholder(ref phFontData, paragraph);
        if (this.TryFromFontData(phFontData))
        {
            return true;
        }

        switch (placeholder.Type)
        {
            case SCPlaceholderType.Title:
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

            case SCPlaceholderType.Text:
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
            colorHexVariant = fontData.ARgbColorModelHex.Val!;
            this.colorType = SCColorType.RGB;
            this.hexColor = colorHexVariant;
            return true;
        }

        if (fontData.ASchemeColor != null)
        {
            colorHexVariant = this.GetHexVariantByScheme(fontData.ASchemeColor.Val!);
            this.colorType = SCColorType.Scheme;
            this.hexColor = colorHexVariant;
            return true;
        }

        if (fontData.ASystemColor != null)
        {
            colorHexVariant = fontData.ASystemColor.LastColor!;
            this.colorType = SCColorType.System;
            this.hexColor = colorHexVariant;
            return true;
        }

        if (fontData.APresetColor != null)
        {
            this.colorType = SCColorType.Preset;
            var coloName = fontData.APresetColor.Val!.Value.ToString();
            this.hexColor = SCColorTranslator.HexFromName(coloName);
            return true;
        }

        return false;
    }

    private string GetHexVariantByScheme(A.SchemeColorValues fontSchemeColor)
    {
        var themeAColorScheme = this.parentSlideMaster.ThemePart.Theme.ThemeElements!.ColorScheme!;
        return fontSchemeColor switch
        {
            A.SchemeColorValues.Dark1 => themeAColorScheme.Dark1Color!.RgbColorModelHex != null
                ? themeAColorScheme.Dark1Color.RgbColorModelHex!.Val!.Value!
                : themeAColorScheme.Dark1Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Light1 => themeAColorScheme.Light1Color!.RgbColorModelHex != null
                ? themeAColorScheme.Light1Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Light1Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Dark2 => themeAColorScheme.Dark2Color!.RgbColorModelHex != null
                ? themeAColorScheme.Dark2Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Dark2Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Light2 => themeAColorScheme.Light2Color!.RgbColorModelHex != null
                ? themeAColorScheme.Light2Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Light2Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent1 => themeAColorScheme.Accent1Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent1Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent1Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent2 => themeAColorScheme.Accent2Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent2Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent2Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent3 => themeAColorScheme.Accent3Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent3Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent3Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent4 => themeAColorScheme.Accent4Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent4Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent4Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent5 => themeAColorScheme.Accent5Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent5Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent5Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Accent6 => themeAColorScheme.Accent6Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent6Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent6Color.SystemColor!.LastColor!.Value!,
            A.SchemeColorValues.Hyperlink => themeAColorScheme.Hyperlink!.RgbColorModelHex != null
                ? themeAColorScheme.Hyperlink.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Hyperlink.SystemColor!.LastColor!.Value!,
            _ => this.GetThemeMappedColor(fontSchemeColor)
        };
    }

    private string GetThemeMappedColor(A.SchemeColorValues fontSchemeColor)
    {
        var slideMasterPColorMap = this.parentSlideMaster.PSlideMaster.ColorMap;
        if (fontSchemeColor == A.SchemeColorValues.Text1)
        {
            return this.GetThemeColorByString(slideMasterPColorMap!.Text1!.ToString() !);
        }

        if (fontSchemeColor == A.SchemeColorValues.Text2)
        {
            return this.GetThemeColorByString(slideMasterPColorMap!.Text2!.ToString() !);
        }

        if (fontSchemeColor == A.SchemeColorValues.Background1)
        {
            return this.GetThemeColorByString(slideMasterPColorMap!.Background1!.ToString() !);
        }

        return this.GetThemeColorByString(slideMasterPColorMap!.Background2!.ToString() !);
    }

    private string GetThemeColorByString(string fontSchemeColor)
    {
        var themeAColorScheme = this.parentSlideMaster.ThemePart.Theme.ThemeElements!.ColorScheme!;
        return fontSchemeColor switch
        {
            "dk1" => themeAColorScheme.Dark1Color!.RgbColorModelHex != null
                ? themeAColorScheme.Dark1Color.RgbColorModelHex!.Val!.Value!
                : themeAColorScheme.Dark1Color.SystemColor!.LastColor!.Value!,
            "lt1" => themeAColorScheme.Light1Color!.RgbColorModelHex != null
                ? themeAColorScheme.Light1Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Light1Color.SystemColor!.LastColor!.Value!,
            "dk2" => themeAColorScheme.Dark2Color!.RgbColorModelHex != null
                ? themeAColorScheme.Dark2Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Dark2Color.SystemColor!.LastColor!.Value!,
            "lt2" => themeAColorScheme.Light2Color!.RgbColorModelHex != null
                ? themeAColorScheme.Light2Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Light2Color.SystemColor!.LastColor!.Value!,
            "accent1" => themeAColorScheme.Accent1Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent1Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent1Color.SystemColor!.LastColor!.Value!,
            "accent2" => themeAColorScheme.Accent2Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent2Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent2Color.SystemColor!.LastColor!.Value!,
            "accent3" => themeAColorScheme.Accent3Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent3Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent3Color.SystemColor!.LastColor!.Value!,
            "accent4" => themeAColorScheme.Accent4Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent4Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent4Color.SystemColor!.LastColor!.Value!,
            "accent5" => themeAColorScheme.Accent5Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent5Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent5Color.SystemColor!.LastColor!.Value!,
            "accent6" => themeAColorScheme.Accent6Color!.RgbColorModelHex != null
                ? themeAColorScheme.Accent6Color.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Accent6Color.SystemColor!.LastColor!.Value!,
            _ => themeAColorScheme.Hyperlink!.RgbColorModelHex != null
                ? themeAColorScheme.Hyperlink.RgbColorModelHex.Val!.Value!
                : themeAColorScheme.Hyperlink.SystemColor!.LastColor!.Value!
        };
    }
}
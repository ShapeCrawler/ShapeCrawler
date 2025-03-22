using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Fonts;

internal readonly record struct IndentFonts
{
    private readonly OpenXmlCompositeElement sdkOpenXmlCompositeElement;

    internal IndentFonts(OpenXmlCompositeElement sdkOpenXmlCompositeElement)
    {
        this.sdkOpenXmlCompositeElement = sdkOpenXmlCompositeElement;
    }

    #region APIs

    internal IndentFont? FontOrNull(int indentLevelFor)
    {
        // Get <a:lvlXpPr> elements, eg. <a:lvl1pPr>, <a:lvl2pPr>
        var lvlParagraphPropertyList = this.sdkOpenXmlCompositeElement.Elements()
            .Where(e => e.LocalName.StartsWith("lvl", StringComparison.Ordinal));

        // Try to find matching font from level-specific paragraph properties
        var indentFont = this.FindFontFromLevelProperties(lvlParagraphPropertyList, indentLevelFor);
        if (indentFont.HasValue)
        {
            return indentFont;
        }

        // Fallback for level 1
        return indentLevelFor == 1 ? this.FindFontFromTextBody() : null;
    }

    private IndentFont? FindFontFromLevelProperties(IEnumerable<OpenXmlElement> lvlParagraphPropertyList, int targetLevel)
    {
        foreach (var textPr in lvlParagraphPropertyList)
        {
            var paragraphLvl = ExtractLevelNumber(textPr.LocalName);
            if (paragraphLvl != targetLevel)
            {
                continue;
            }

            var aDefRPr = textPr.GetFirstChild<A.DefaultRunProperties>();
            if (aDefRPr == null)
            {
                continue;
            }

            return CreateIndentFont(aDefRPr);
        }

        return null;
    }

    private static int ExtractLevelNumber(string localName)
    {
#if NETSTANDARD2_0
        return int.Parse(
            localName[3].ToString(System.Globalization.CultureInfo.CurrentCulture),
            System.Globalization.CultureInfo.CurrentCulture);
#else
        var nameSpan = localName.AsSpan();
        var level = nameSpan.Slice(3, 1); // the fourth character contains level number, eg. "lvl1pPr -> 1, lvl2pPr -> 2, etc."
        return int.Parse(level, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.CurrentCulture);
#endif
    }

    private static IndentFont CreateIndentFont(A.DefaultRunProperties aDefRPr)
    {
        var fontSize = aDefRPr.FontSize?.Value;
        var isBold = aDefRPr.Bold?.Value;
        var isItalic = aDefRPr.Italic;
        var aLatinFont = aDefRPr.GetFirstChild<A.LatinFont>();

        // Extract color properties
        var (aRgbColorModelHex, aSchemeColor, aSystemColor, aPresetColor) = ExtractColorProperties(aDefRPr);

        return new IndentFont
        {
            Size = fontSize,
            ALatinFont = aLatinFont,
            IsBold = isBold,
            IsItalic = isItalic,
            ARgbColorModelHex = aRgbColorModelHex,
            ASchemeColor = aSchemeColor,
            ASystemColor = aSystemColor,
            APresetColor = aPresetColor
        };
    }

    private static (A.RgbColorModelHex?, A.SchemeColor?, A.SystemColor?, A.PresetColor?) ExtractColorProperties(A.DefaultRunProperties aDefRPr)
    {
        // Try get color from <a:solidFill>
        var aSolidFill = aDefRPr.SdkASolidFill();
        if (aSolidFill != null)
        {
            return (aSolidFill.RgbColorModelHex, aSolidFill.SchemeColor, aSolidFill.SystemColor, aSolidFill.PresetColor);
        }

        // Try get color from gradient fill
        var aGradientStop = aDefRPr.GetFirstChild<A.GradientFill>()?.GradientStopList?
            .GetFirstChild<A.GradientStop>();
            
        return (aGradientStop?.RgbColorModelHex, aGradientStop?.SchemeColor, 
                aGradientStop?.SystemColor, aGradientStop?.PresetColor);
    }

    private IndentFont? FindFontFromTextBody()
    {
        if (this.sdkOpenXmlCompositeElement.Parent is not P.TextBody pTextBody)
        {
            return null;
        }

        var endParaRunPrFs = pTextBody.GetFirstChild<A.Paragraph>()?
            .GetFirstChild<A.EndParagraphRunProperties>()?.FontSize;
            
        if (endParaRunPrFs is null)
        {
            return null;
        }

        return new IndentFont
        {
            Size = endParaRunPrFs
        };
    }

    internal ColorType? ColorType(int indentLevel)
    {
        var indentFont = this.FontOrNull(indentLevel);
        if (indentFont is null)
        {
            return null;
        }

        if (indentFont.Value.ARgbColorModelHex != null)
        {
            return ShapeCrawler.ColorType.RGB;
        }

        if (indentFont.Value.ASchemeColor != null)
        {
            return ShapeCrawler.ColorType.Theme;
        }

        if (indentFont.Value.ASystemColor != null)
        {
            return ShapeCrawler.ColorType.Standard;
        }

        if (indentFont.Value.APresetColor != null)
        {
            return ShapeCrawler.ColorType.Preset;
        }

        return null;
    }

    internal bool? BoldFlagOrNull(int indentLevel)
    {
        var indentFont = this.FontOrNull(indentLevel);

        return indentFont?.IsBold;
    }

    internal A.LatinFont? ALatinFontOrNull(int indentLevel)
    {
        var indentFont = this.FontOrNull(indentLevel);

        return indentFont?.ALatinFont;
    }

    #endregion APIs
}
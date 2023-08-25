using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Extensions;
using ShapeCrawler.Fonts;
using ShapeCrawler.Services;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal sealed class SlideFontColor : IFontColor
{
    private readonly A.Text aText;
    private bool initialized;
    private string? hexColor;
    private SCColorType colorType;
    private readonly SlidePart sdkSlidePart;

    internal SlideFontColor(SlidePart sdkSlidePart, A.Text aText)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.aText = aText;
    }

    public SCColorType ColorType => this.ParseColorType();

    public string ColorHex => this.ParseHexColor();

    public void SetColorByHex(string hex)
    {
        var aTextContainer = this.aText.Parent!;
        var aRunProperties = aTextContainer.GetFirstChild<A.RunProperties>() ?? aTextContainer.AddRunProperties();

        var aSolidFill = aRunProperties.GetASolidFill();
        aSolidFill?.Remove();

        // All hex values are expected to be without hashtag.
        hex = hex.StartsWith("#", System.StringComparison.Ordinal) ? hex.Substring(1) : hex; // to skip '#'
        var rgbColorModelHex = new A.RgbColorModelHex { Val = hex };
        aSolidFill = new A.SolidFill();
        aSolidFill.Append(rgbColorModelHex);
        aRunProperties.Append(aSolidFill);
    }

    private SCColorType ParseColorType()
    {
        if (!this.initialized)
        {
            this.InitializeColor();
        }

        return this.colorType;
    }

    private string ParseHexColor()
    {
        if (!this.initialized)
        {
            this.InitializeColor();
        }

        return this.hexColor!;
    }

    private void InitializeColor()
    {
        var pSlideMaster = this.sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster;
        this.initialized = true;
        var aSolidFill = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.GetASolidFill();
        if (aSolidFill != null)
        {
            var typeAndHex = HexParser.FromSolidFill(aSolidFill, pSlideMaster);
            this.colorType = typeAndHex.Item1;
            this.hexColor = typeAndHex.Item2;
        }
        else
        {
            if (this.TryFromTextBody())
            {
                return;
            }

            if (this.TryFromShapeFontReference())
            {
                return;
            }
            
            var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
            var paraLevel = new AParagraphWrap(aParagraph).IndentLevel();
            if (this.TryFromPlaceholder(paraLevel))
            {
                return;
            }
            
            var pSlideMasterWrap =
                new PSlideMasterWrap(pSlideMaster);
            ParagraphLevelFont? masterBodyStyleParaLevelFont = pSlideMasterWrap.BodyStyleParagraphLevelFontOrNull(paraLevel);
            if (masterBodyStyleParaLevelFont != null && this.TryFromFontData(masterBodyStyleParaLevelFont))
            {
                return;
            }

            PresentationPart sdkPresentationPart = (PresentationPart)this.sdkSlidePart.OpenXmlPackage.RootPart!;
            var pPresentationWrap = new PPresentationWrap(sdkPresentationPart.Presentation);
            ParagraphLevelFont? presentationParaLevelFont = this.parentFont.Presentation().FontDataOrNullForParagraphLevel(paraLevel);
            // Presentation level
            string hexColor;
            if (presentationParaLevelFont != null)
            {
                hexColor = this.HexColorByScheme(presentationParaLevelFont.ASchemeColor!.Val!);
                this.colorType = SCColorType.Theme;
                this.hexColor = hexColor;
                return;
            }

            // Get default
            hexColor = this.GetThemeMappedColor(A.SchemeColorValues.Text1);
            this.colorType = SCColorType.Theme;
            this.hexColor = hexColor;
        }
    }

    private bool TryFromTextBody()
    {
        var paraLvlToFontData = FontDataParser.FromCompositeElement(this.textBodyListStyle!);
        if (!paraLvlToFontData.TryGetValue(this.parentFont.ParagraphLevel(), out var txBodyFontData))
        {
            return false;
        }

        return this.TryFromFontData(txBodyFontData);
    }

    private bool TryFromShapeFontReference()
    {
        SlideAutoShape autoShape = this.parentFont.SlideAutoShape();
        if (this.textFrameContainer is SCShape parentShape)
        {
            var parentPShape = (DocumentFormat.OpenXml.Presentation.Shape)parentShape.PShapeTreeChild;
            if (parentPShape.ShapeStyle == null)
            {
                return false;
            }

            var aFontReference = parentPShape.ShapeStyle.FontReference!;
            var fontReferenceFontData = new ParagraphLevelFont
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
        if (this.textFrameContainer.AutoShape.Placeholder is not SCPlaceholder placeholder)
        {
            return false;
        }

        var phFontData = new ParagraphLevelFont();
        FontDataParser.GetFontDataFromPlaceholder(ref phFontData, this.paragraph);
        if (this.TryFromFontData(phFontData))
        {
            return true;
        }

        switch (placeholder.Type)
        {
            case SCPlaceholderType.Title:
            {
                Dictionary<int, ParagraphLevelFont> titleParaLvlToFontData = this.parentSlideMaster.TitleParaLvlToFontData;
                ParagraphLevelFont masterTitleParagraphLevelFont = titleParaLvlToFontData.ContainsKey(paragraphLevel)
                    ? titleParaLvlToFontData[paragraphLevel]
                    : titleParaLvlToFontData[1];
                if (this.TryFromFontData(masterTitleParagraphLevelFont))
                {
                    return true;
                }

                break;
            }

            case SCPlaceholderType.Text:
            {
                Dictionary<int, ParagraphLevelFont> bodyParaLvlToFontData = this.parentSlideMaster.BodyParaLvlToFontData;
                ParagraphLevelFont masterBodyParagraphLevelFont = bodyParaLvlToFontData[paragraphLevel];
                if (this.TryFromFontData(masterBodyParagraphLevelFont))
                {
                    return true;
                }

                break;
            }
        }

        return false;
    }

    private bool TryFromFontData(ParagraphLevelFont paragraphLevelFont)
    {
        string colorHexVariant;
        if (paragraphLevelFont.ARgbColorModelHex != null)
        {
            colorHexVariant = paragraphLevelFont.ARgbColorModelHex.Val!;
            this.colorType = SCColorType.RGB;
            this.hexColor = colorHexVariant;
            return true;
        }

        if (paragraphLevelFont.ASchemeColor != null)
        {
            colorHexVariant = this.HexColorByScheme(paragraphLevelFont.ASchemeColor.Val!);
            this.colorType = SCColorType.Theme;
            this.hexColor = colorHexVariant;
            return true;
        }

        if (paragraphLevelFont.ASystemColor != null)
        {
            colorHexVariant = paragraphLevelFont.ASystemColor.LastColor!;
            this.colorType = SCColorType.Standard;
            this.hexColor = colorHexVariant;
            return true;
        }

        if (paragraphLevelFont.APresetColor != null)
        {
            this.colorType = SCColorType.Preset;
            var coloName = paragraphLevelFont.APresetColor.Val!.Value.ToString();
            this.hexColor = SCColorTranslator.HexFromName(coloName);
            return true;
        }

        return false;
    }

    private string HexColorByScheme(A.SchemeColorValues fontSchemeColor)
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

internal sealed record PPresentationWrap
{
    private readonly Presentation pPresentation;

    internal PPresentationWrap(Presentation pPresentation)
    {
        this.pPresentation = pPresentation;
    }
}
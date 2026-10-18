using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Colors;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Fonts;

internal sealed class TextPortionFont(
    IFontSize fontSize,
    Lazy<FontColor> fontColor,
    ThemeFontScheme themeFontScheme,
    A.Text aText) : ITextPortionFont
{
    public decimal Size
    {
        get => fontSize.Size;
        set => fontSize.Size = value;
    }

    public string? LatinName
    {
        get
        {
            if (this.ALatinFont().Typeface == "+mj-lt")
            {
                return themeFontScheme.MajorLatinFont();
            }

            return this.ALatinFont().Typeface!;
        }
        set => this.UpdateLatinName(value!);
    }

    public string EastAsianName
    {
        get => new SCAText(aText).EastAsianName();
        set => new SCAText(aText).UpdateEastAsianName(value);
    }

    public bool IsBold
    {
        get => new TextStyleFlags(aText).Bold;
        set => this.UpdateBoldFlag(value);
    }

    public bool IsItalic
    {
        get => new TextStyleFlags(aText).Italic;
        set => this.SetItalicFlag(value);
    }

    public IFontColor Color => fontColor.Value;

    public A.TextUnderlineValues Underline
    {
        get
        {
            var aRunPr = aText.Parent!.GetFirstChild<A.RunProperties>();
            return aRunPr?.Underline?.Value ?? A.TextUnderlineValues.None;
        }

        set
        {
            var aRunPr = aText.Parent!.GetFirstChild<A.RunProperties>();
            if (aRunPr != null)
            {
                aRunPr.Underline = new EnumValue<A.TextUnderlineValues>(value);
            }
            else
            {
                var aEndParaRPr =
                    aText.Parent.NextSibling<A.EndParagraphRunProperties>();
                if (aEndParaRPr != null)
                {
                    aEndParaRPr.Underline = new EnumValue<A.TextUnderlineValues>(value);
                }
                else
                {
                    var runProp = aText.Parent.AddRunProperties();
                    runProp.Underline = new EnumValue<A.TextUnderlineValues>(value);
                }
            }
        }
    }

    public int OffsetEffect
    {
        get => this.GetOffsetEffect();
        set => this.SetOffset(value);
    }

    private void SetOffset(int value)
    {
        var aRunProperties = aText.Parent!.GetFirstChild<A.RunProperties>();
        Int32Value int32Value = value * 1000;
        if (aRunProperties is not null)
        {
            aRunProperties.Baseline = int32Value;
        }
        else
        {
            var aEndParaRPr = aText.Parent.NextSibling<A.EndParagraphRunProperties>();
            if (aEndParaRPr != null)
            {
                aEndParaRPr.Baseline = int32Value;
            }
            else
            {
                aRunProperties = new A.RunProperties { Baseline = int32Value };
                aText.Parent.InsertAt(aRunProperties, 0); // append to <a:r>
            }
        }
    }

    private int GetOffsetEffect()
    {
        var aRunProperties = aText.Parent!.GetFirstChild<A.RunProperties>();
        if (aRunProperties is not null &&
            aRunProperties.Baseline is not null)
        {
            return aRunProperties.Baseline.Value / 1000;
        }

        var aEndParaRPr = aText.Parent.NextSibling<A.EndParagraphRunProperties>();
        if (aEndParaRPr is not null)
        {
            return aEndParaRPr.Baseline! / 1000;
        }

        return 0;
    }

    private A.LatinFont ALatinFont()
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var aRunProperties = aText.Parent!.GetFirstChild<A.RunProperties>();
        var aLatinFont = aRunProperties?.GetFirstChild<A.LatinFont>();

        if (aLatinFont != null)
        {
            return aLatinFont;
        }

        aLatinFont = new ReferencedFont(new ReferencedFontColor(aText), aText).ALatinFontOrNull();
        if (aLatinFont != null)
        {
            return aLatinFont;
        }

        return new ThemeFontScheme(openXmlPart).MinorLatinFont();
    }

    private void UpdateBoldFlag(bool value)
    {
        var properties = aText.Parent!.GetFirstChild<A.RunProperties>() ??
            aText.Parent.PrependChild(new A.RunProperties());
        properties.Bold = value;
    }

    private void SetItalicFlag(bool isItalic)
    {
        var properties = aText.Parent!.GetFirstChild<A.RunProperties>() ??
            aText.Parent.PrependChild(new A.RunProperties());
        properties.Italic = isItalic;
    }

    private void UpdateLatinName(string latinFont)
    {
        var aRunProperties = aText.Parent!.GetFirstChild<A.RunProperties>();

        A.TextCharacterPropertiesType aCurrentProperties;

        if (aRunProperties is not null)
        {
            aCurrentProperties = aRunProperties;
        }
        else
        {
            var aEndParaRunProperties = aText.Parent!.GetFirstChild<A.EndParagraphRunProperties>();

            if (aEndParaRunProperties is not null)
            {
                aCurrentProperties = aEndParaRunProperties;
            }
            else
            {
                aCurrentProperties = aText.Parent!.AddRunProperties();
            }
        }

        var aLatinFont = aCurrentProperties.GetFirstChild<A.LatinFont>();

        if (aLatinFont is null)
        {
            aLatinFont = new A.LatinFont();
            aCurrentProperties.Append(aLatinFont); // to avoid ignoring color
        }

        aLatinFont.Typeface = latinFont;
    }
}
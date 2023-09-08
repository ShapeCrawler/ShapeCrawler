using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using ShapeCrawler.Wrappers;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Fonts;

internal sealed class SlideTextPortionFont : ITextPortionFont
{
    private readonly A.Text aText;
    private readonly Lazy<SlideFontColor> fontColor;
    private readonly ResetableLazy<A.LatinFont> latinFont;
    private readonly IFontSize size;
    private readonly ThemeFontScheme themeFontScheme;
    private readonly SlidePart sdkSlidePart;
    private readonly SdkSlideAText sdkAText;

    internal SlideTextPortionFont(
        SlidePart sdkSlidePart,
        A.Text aText,
        IFontSize size)
        : this(
            sdkSlidePart,
            aText,
            size,
            new ThemeFontScheme(sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .FontScheme!)
        )
    {
    }

    private SlideTextPortionFont(
        SlidePart sdkSlidePart,
        A.Text aText,
        IFontSize size,
        ThemeFontScheme themeFontScheme)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.aText = aText;
        this.latinFont = new ResetableLazy<A.LatinFont>(this.ParseALatinFont);
        this.fontColor = new Lazy<SlideFontColor>(() => new SlideFontColor(this.sdkSlidePart, this.aText));
        this.size = size;
        this.themeFontScheme = themeFontScheme;
        this.sdkAText = new SdkSlideAText(sdkSlidePart, aText);
    }

    public int Size
    {
        get => this.size.Size();
        set => this.size.Update(value);
    }

    public string? LatinName
    {
        get => this.ParseLatinName();
        set => this.SetLatinName(value!);
    }

    public string EastAsianName
    {
        get => this.sdkAText.EastAsianName();
        set => this.sdkAText.UpdateEastAsianName(value);
    }

    public bool IsBold
    {
        get => this.GetBoldFlag();
        set => this.SetBoldFlag(value);
    }

    public bool IsItalic
    {
        get => this.GetItalicFlag();
        set => this.SetItalicFlag(value);
    }

    public IFontColor Color => this.fontColor.Value;

    public A.TextUnderlineValues Underline
    {
        get
        {
            var aRunPr = this.aText.Parent!.GetFirstChild<A.RunProperties>();
            return aRunPr?.Underline?.Value ?? A.TextUnderlineValues.None;
        }

        set
        {
            var aRunPr = this.aText.Parent!.GetFirstChild<A.RunProperties>();
            if (aRunPr != null)
            {
                aRunPr.Underline = new EnumValue<A.TextUnderlineValues>(value);
            }
            else
            {
                var aEndParaRPr =
                    this.aText.Parent.NextSibling<A.EndParagraphRunProperties>();
                if (aEndParaRPr != null)
                {
                    aEndParaRPr.Underline = new EnumValue<A.TextUnderlineValues>(value);
                }
                else
                {
                    var runProp = this.aText.Parent.AddRunProperties();
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
        var aRunProperties = this.aText.Parent!.GetFirstChild<A.RunProperties>();
        Int32Value int32Value = value * 1000;
        if (aRunProperties is not null)
        {
            aRunProperties.Baseline = int32Value;
        }
        else
        {
            var aEndParaRPr = this.aText.Parent.NextSibling<A.EndParagraphRunProperties>();
            if (aEndParaRPr != null)
            {
                aEndParaRPr.Baseline = int32Value;
            }
            else
            {
                aRunProperties = new A.RunProperties { Baseline = int32Value };
                this.aText.Parent.InsertAt(aRunProperties, 0); // append to <a:r>
            }
        }
    }

    private int GetOffsetEffect()
    {
        var aRunProperties = this.aText.Parent!.GetFirstChild<A.RunProperties>();
        if (aRunProperties is not null &&
            aRunProperties.Baseline is not null)
        {
            return aRunProperties.Baseline.Value / 1000;
        }

        var aEndParaRPr = this.aText.Parent.NextSibling<A.EndParagraphRunProperties>();
        if (aEndParaRPr is not null)
        {
            return aEndParaRPr.Baseline! / 1000;
        }

        return 0;
    }

    private string ParseLatinName()
    {
        if (this.latinFont.Value.Typeface == "+mj-lt")
        {
            return this.themeFontScheme.MajorLatinFont();
        }

        return this.latinFont.Value.Typeface!;
    }

    private string ParseEastAsianName()
    {
        var aEastAsianFont = this.AEastAsianFont();
        if (aEastAsianFont.Typeface == "+mj-ea")
        {
            return this.themeFontScheme.MajorEastAsianFont();
        }

        return aEastAsianFont.Typeface!;
    }

    private A.EastAsianFont AEastAsianFont()
    {
        var aEastAsianFont = this.aText.Parent!.GetFirstChild<A.RunProperties>()
            ?.GetFirstChild<A.EastAsianFont>();

        if (aEastAsianFont != null)
        {
            return aEastAsianFont;
        }

        throw new Exception("TODO: implement");
    }

    private A.LatinFont ParseALatinFont()
    {
        var aRunProperties = this.aText.Parent!.GetFirstChild<A.RunProperties>();
        var aLatinFont = aRunProperties?.GetFirstChild<A.LatinFont>();

        if (aLatinFont != null)
        {
            return aLatinFont;
        }

        throw new Exception("TODO: implement");
    }

    private bool GetBoldFlag()
    {
        var aRunProperties = this.aText.Parent!.GetFirstChild<A.RunProperties>();
        if (aRunProperties == null)
        {
            return false;
        }

        if (aRunProperties.Bold is not null && aRunProperties.Bold == true)
        {
            return true;
        }

        IndentFont phIndentFont = new();
        if (phIndentFont.IsBold is not null)
        {
            return phIndentFont.IsBold.Value;
        }

        return false;
    }

    private bool GetItalicFlag()
    {
        var aRunPr = this.aText.Parent!.GetFirstChild<A.RunProperties>();
        if (aRunPr == null)
        {
            return false;
        }

        if (aRunPr.Italic is not null && aRunPr.Italic == true)
        {
            return true;
        }

        IndentFont phIndentFont = new();
        if (phIndentFont.IsItalic is not null)
        {
            return phIndentFont.IsItalic.Value;
        }

        return false;
    }

    private void SetBoldFlag(bool value)
    {
        var aRunPr = this.aText.Parent!.GetFirstChild<A.RunProperties>();
        if (aRunPr != null)
        {
            aRunPr.Bold = new BooleanValue(value);
        }
        else
        {
            IndentFont phIndentFont = new();
            if (phIndentFont.IsBold is not null)
            {
                phIndentFont.IsBold = new BooleanValue(value);
            }
            else
            {
                var aEndParaRPr =
                    this.aText.Parent.NextSibling<A.EndParagraphRunProperties>();
                if (aEndParaRPr != null)
                {
                    aEndParaRPr.Bold = new BooleanValue(value);
                }
                else
                {
                    aRunPr = new A.RunProperties { Bold = new BooleanValue(value) };
                    this.aText.Parent.InsertAt(aRunPr, 0); // append to <a:r>
                }
            }
        }
    }

    private void SetItalicFlag(bool isItalic)
    {
        var aTextParent = this.aText.Parent!;
        var aRunPr = aTextParent.GetFirstChild<A.RunProperties>();
        if (aRunPr != null)
        {
            aRunPr.Italic = new BooleanValue(isItalic);
        }
        else
        {
            var aEndParaRPr = aTextParent.NextSibling<A.EndParagraphRunProperties>();
            if (aEndParaRPr != null)
            {
                aEndParaRPr.Italic = new BooleanValue(isItalic);
            }
            else
            {
                aTextParent.AddRunProperties(isItalic);
            }
        }
    }

    private void SetLatinName(string latinFont)
    {
        var aLatinFont = this.latinFont.Value;
        aLatinFont.Typeface = latinFont;
        this.latinFont.Reset();
    }

    private void UpdateEastAsianName(string eastAsianFont)
    {
        var aEastAsianFont = this.AEastAsianFont();
        aEastAsianFont.Typeface = eastAsianFont;
    }
}
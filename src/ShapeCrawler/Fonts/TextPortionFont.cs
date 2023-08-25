
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Services;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal sealed class TextPortionFont : ITextPortionFont
{
    private readonly A.Text aText;
    private readonly A.FontScheme aFontScheme;
    private readonly Lazy<SlideFontColor> fontColor;
    private readonly ResetableLazy<A.LatinFont> latinFont;
    private readonly IFontSize size;
    private readonly SCParagraphTextPortion parentParagraphTextPortion;

    internal TextPortionFont(
        A.Text aText, 
        ThemeFontScheme themeFontScheme,
        IFontSize size,
        SCParagraphTextPortion parentParagraphTextPortion)
    {
        this.parentParagraphTextPortion = parentParagraphTextPortion;
        this.aText = aText;
        this.latinFont = new ResetableLazy<A.LatinFont>(this.ParseALatinFont);
        
        A.ListStyle textBodyListStyle = this.parentParagraphTextPortion.ATextBodyListStyle();
        this.fontColor = new Lazy<SlideFontColor>(() => new SlideFontColor(this.aText, textBodyListStyle));
        this.aFontScheme = themeFontScheme.AFontScheme;
        this.size = size;
    }

    public int Size
    {
        get => this.size.Size();
        set => this.size.Update(value);
    }
    
    public string? LatinName
    {
        get => this.GetLatinName();
        set => this.SetLatinName(value!);
    }

    public string? EastAsianName
    {
        get => this.GetEastAsianName();
        set => this.SetEastAsianName(value!);
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

    public bool CanChange()
    {
        var placeholder = this.textFrameContainer.AutoShape.Placeholder;

        return placeholder is null or { Type: SCPlaceholderType.Text };
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

    private string GetLatinName()
    {
        if (this.latinFont.Value.Typeface == "+mj-lt")
        {
            return this.aFontScheme.MajorFont!.LatinFont!.Typeface!;
        }

        return this.latinFont.Value.Typeface!;
    }

    private string GetEastAsianName()
    {
        var aEastAsianFont = this.GetAEastAsianFont();
        if (aEastAsianFont.Typeface == "+mj-ea")
        {
            return this.aFontScheme.MajorFont!.EastAsianFont!.Typeface!;
        }

        return aEastAsianFont.Typeface!;
    }

    private A.EastAsianFont GetAEastAsianFont()
    {
        var aEastAsianFont = this.aText.Parent!.GetFirstChild<A.RunProperties>()
            ?.GetFirstChild<A.EastAsianFont>();

        if (aEastAsianFont != null)
        {
            return aEastAsianFont;
        }

        var phFontData = FontDataParser.FromPlaceholder(this.paragraph);

        return phFontData.AEastAsianFont ?? this.aFontScheme.MinorFont!.EastAsianFont!;
    }

    private A.LatinFont ParseALatinFont()
    {
        var aRunProperties = this.aText.Parent!.GetFirstChild<A.RunProperties>();
        var aLatinFont = aRunProperties?.GetFirstChild<A.LatinFont>();

        if (aLatinFont != null)
        {
            return aLatinFont;
        }

        var phFontData = FontDataParser.FromPlaceholder(this.paragraph);
        return phFontData.ALatinFont ?? this.aFontScheme.MinorFont!.LatinFont!;
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

        FontData phFontData = new();
        FontDataParser.GetFontDataFromPlaceholder(ref phFontData, this.paragraph);
        if (phFontData.IsBold is not null)
        {
            return phFontData.IsBold.Value;
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

        FontData phFontData = new();
        FontDataParser.GetFontDataFromPlaceholder(ref phFontData, this.paragraph);
        if (phFontData.IsItalic is not null)
        {
            return phFontData.IsItalic.Value;
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
            FontData phFontData = new();
            FontDataParser.GetFontDataFromPlaceholder(ref phFontData, this.paragraph);
            if (phFontData.IsBold is not null)
            {
                phFontData.IsBold = new BooleanValue(value);
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

    private void SetEastAsianName(string eastAsianFont)
    {
        var aEastAsianFont = this.GetAEastAsianFont();
        aEastAsianFont.Typeface = eastAsianFont;
    }

    internal SlideMaster SlideMaster()
    {
        return this.parentParagraphTextPortion.SlideMaster();
    }

    internal int ParagraphLevel()
    {
        return this.parentParagraphTextPortion.ParagraphLevel();
    }

    internal PresentationCore Presentation()
    {
        return this.parentParagraphTextPortion.Presentation();
    }

    internal SlideAutoShape SlideAutoShape()
    {
        return this.parentParagraphTextPortion.SlideAutoShape();
    }
}
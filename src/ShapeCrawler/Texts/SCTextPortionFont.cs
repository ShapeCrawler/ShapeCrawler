using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Constants;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Services;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;

namespace ShapeCrawler.Texts;

internal sealed class SCTextPortionFont : ITextPortionFont
{
    private readonly DocumentFormat.OpenXml.Drawing.Text aText;
    private readonly DocumentFormat.OpenXml.Drawing.FontScheme aFontScheme;
    private readonly Lazy<SCFontColor> colorFormat;
    private readonly ResetableLazy<DocumentFormat.OpenXml.Drawing.LatinFont> latinFont;
    private readonly ResetableLazy<int> size;
    private readonly ITextFrameContainer textFrameContainer;
    private readonly SCParagraph paragraph;

    internal SCTextPortionFont(DocumentFormat.OpenXml.Drawing.Text aText, IPortion portion, ITextFrameContainer textFrameContainer, SCParagraph paragraph)
    {
        this.aText = aText;
        this.paragraph = paragraph;
        this.size = new ResetableLazy<int>(this.GetSize);
        this.latinFont = new ResetableLazy<DocumentFormat.OpenXml.Drawing.LatinFont>(this.GetALatinFont);
        this.colorFormat = new Lazy<SCFontColor>(() => new SCFontColor(this, textFrameContainer, paragraph, this.aText));
        this.ParentPortion = portion;
        SCShape shape;
        this.textFrameContainer = textFrameContainer;
        if (textFrameContainer is SCCell cell)
        {
            shape = cell.SCShape;
        }
        else
        {
            shape = (SCShape)textFrameContainer;
        }

        var themeFontScheme = (ThemeFontScheme)shape.SlideMasterInternal.Theme.FontScheme; 
        this.aFontScheme = themeFontScheme.AFontScheme;
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

    public int Size
    {
        get => this.size.Value;
        set => this.SetFontSize(value);
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

    public IFontColor Color => this.colorFormat.Value;

    public DocumentFormat.OpenXml.Drawing.TextUnderlineValues Underline
    {
        get
        {
            var aRunPr = this.aText.Parent!.GetFirstChild<DocumentFormat.OpenXml.Drawing.RunProperties>();
            return aRunPr?.Underline?.Value ?? DocumentFormat.OpenXml.Drawing.TextUnderlineValues.None;
        }

        set
        {
            var aRunPr = this.aText.Parent!.GetFirstChild<DocumentFormat.OpenXml.Drawing.RunProperties>();
            if (aRunPr != null)
            {
                aRunPr.Underline = new EnumValue<DocumentFormat.OpenXml.Drawing.TextUnderlineValues>(value);
            }
            else
            {
                var aEndParaRPr = this.aText.Parent.NextSibling<DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties>();
                if (aEndParaRPr != null)
                {
                    aEndParaRPr.Underline = new EnumValue<DocumentFormat.OpenXml.Drawing.TextUnderlineValues>(value);
                }
                else
                {
                    var runProp = this.aText.Parent.AddRunProperties();
                    runProp.Underline = new EnumValue<DocumentFormat.OpenXml.Drawing.TextUnderlineValues>(value);
                }
            }
        }
    }

    public int OffsetEffect
    {
        get => this.GetOffsetEffect();
        set => this.SetOffset(value);
    }

    internal IPortion ParentPortion { get; }

    public bool CanChange()
    {
        var placeholder = this.textFrameContainer.SCShape.Placeholder;

        return placeholder is null or { Type: SCPlaceholderType.Text };
    }

    private static bool TryFromPlaceholder(SCShape scShape, int paraLevel, out int i)
    {
        i = -1;
        var placeholder = scShape.Placeholder as SCPlaceholder;
        var referencedShape = placeholder?.ReferencedShape.Value as SCAutoShape;
        var fontDataPlaceholder = new FontData();
        if (referencedShape != null)
        {
            referencedShape.FillFontData(paraLevel, ref fontDataPlaceholder);
            if (fontDataPlaceholder.FontSize is not null)
            {
                {
                    i = fontDataPlaceholder.FontSize / 100;
                    return true;
                }
            }
        }

        var slideMaster = scShape.SlideMasterInternal;
        if (placeholder?.Type == SCPlaceholderType.Title)
        {
            var pTextStyles = slideMaster.PSlideMaster.TextStyles!;
            var titleFontSize = pTextStyles.TitleStyle!.Level1ParagraphProperties!
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.DefaultRunProperties>() !.FontSize!.Value;
            i = titleFontSize / 100;
            return true;
        }

        if (slideMaster.TryGetFontSizeFromBody(paraLevel, out var fontSizeBody))
        {
            {
                i = fontSizeBody / 100;
                return true;
            }
        }

        if (slideMaster.TryGetFontSizeFromOther(paraLevel, out var fontSizeOther))
        {
            {
                i = fontSizeOther / 100;
                return true;
            }
        }

        return false;
    }
    
    private void SetOffset(int value)
    {
        var aRunProperties = this.aText.Parent!.GetFirstChild<DocumentFormat.OpenXml.Drawing.RunProperties>();
        Int32Value int32Value = value * 1000;
        if (aRunProperties is not null)
        {
            aRunProperties.Baseline = int32Value;
        }
        else
        {
            var aEndParaRPr = this.aText.Parent.NextSibling<DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties>();
            if (aEndParaRPr != null)
            {
                aEndParaRPr.Baseline = int32Value;
            }
            else
            {
                aRunProperties = new DocumentFormat.OpenXml.Drawing.RunProperties { Baseline = int32Value };
                this.aText.Parent.InsertAt(aRunProperties, 0); // append to <a:r>
            }
        }
    }

    private int GetOffsetEffect()
    {
        var aRunProperties = this.aText.Parent!.GetFirstChild<DocumentFormat.OpenXml.Drawing.RunProperties>();
        if (aRunProperties is not null &&
            aRunProperties.Baseline is not null)
        {
            return aRunProperties.Baseline.Value / 1000;
        }

        var aEndParaRPr = this.aText.Parent.NextSibling<DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties>();
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
    
    private string? GetEastAsianName()
    {
        var aEastAsianFont = this.GetAEastAsianFont();
        if (aEastAsianFont.Typeface == "+mj-ea")
        {
            return this.aFontScheme.MajorFont!.EastAsianFont!.Typeface!;
        }

        return aEastAsianFont.Typeface!;
    }

    private DocumentFormat.OpenXml.Drawing.EastAsianFont GetAEastAsianFont()
    {
        var aEastAsianFont = this.aText.Parent!.GetFirstChild<DocumentFormat.OpenXml.Drawing.RunProperties>()?.GetFirstChild<DocumentFormat.OpenXml.Drawing.EastAsianFont>();

        if (aEastAsianFont != null)
        {
            return aEastAsianFont;
        }

        var phFontData = FontDataParser.FromPlaceholder(this.paragraph);
        
        return phFontData.AEastAsianFont ?? this.aFontScheme.MinorFont!.EastAsianFont!;
    }
    
    private DocumentFormat.OpenXml.Drawing.LatinFont GetALatinFont()
    {
        var aRunProperties = this.aText.Parent!.GetFirstChild<DocumentFormat.OpenXml.Drawing.RunProperties>();
        var aLatinFont = aRunProperties?.GetFirstChild<DocumentFormat.OpenXml.Drawing.LatinFont>();

        if (aLatinFont != null)
        {
            return aLatinFont;
        }

        var phFontData = FontDataParser.FromPlaceholder(this.paragraph);
        return phFontData.ALatinFont ?? this.aFontScheme.MinorFont!.LatinFont!;
    }

    private int GetSize()
    {
        var fontSize = this.aText.Parent!.GetFirstChild<DocumentFormat.OpenXml.Drawing.RunProperties>()?.FontSize?.Value;
        if (fontSize != null)
        {
            return fontSize.Value / 100;
        }

        var textFrameContainer = this.paragraph.ParentTextFrame.TextFrameContainer;
        var paraLevel = this.paragraph.Level;

        if (textFrameContainer is SCShape { Placeholder: { } } shape)
        {
            if (TryFromPlaceholder(shape, paraLevel, out var sizeFromPlaceholder))
            {
                return sizeFromPlaceholder;
            }
        }

        var sldStructureCore = (SlideStructure)textFrameContainer.SCShape.SlideStructure;
        var pres = sldStructureCore.PresentationInternal;
        if (pres.ParaLvlToFontData.TryGetValue(paraLevel, out var fontData))
        {
            if (fontData.FontSize is not null)
            {
                return fontData.FontSize / 100;
            }
        }

        return SCConstants.DefaultFontSize;
    }

    private bool GetBoldFlag()
    {
        var aRunProperties = this.aText.Parent!.GetFirstChild<DocumentFormat.OpenXml.Drawing.RunProperties>();
        if (aRunProperties == null)
        {
            return false;
        }

        if (aRunProperties.Bold is not null && aRunProperties.Bold == true)
        {
            return true;
        }

        FontData phFontData = new ();
        FontDataParser.GetFontDataFromPlaceholder(ref phFontData, this.paragraph);
        if (phFontData.IsBold is not null)
        {
            return phFontData.IsBold.Value;
        }

        return false;
    }

    private bool GetItalicFlag()
    {
        var aRunPr = this.aText.Parent!.GetFirstChild<DocumentFormat.OpenXml.Drawing.RunProperties>();
        if (aRunPr == null)
        {
            return false;
        }

        if (aRunPr.Italic is not null && aRunPr.Italic == true)
        {
            return true;
        }

        FontData phFontData = new ();
        FontDataParser.GetFontDataFromPlaceholder(ref phFontData, this.paragraph);
        if (phFontData.IsItalic is not null)
        {
            return phFontData.IsItalic.Value;
        }

        return false;
    }

    private void SetBoldFlag(bool value)
    {
        var aRunPr = this.aText.Parent!.GetFirstChild<DocumentFormat.OpenXml.Drawing.RunProperties>();
        if (aRunPr != null)
        {
            aRunPr.Bold = new BooleanValue(value);
        }
        else
        {
            FontData phFontData = new ();
            FontDataParser.GetFontDataFromPlaceholder(ref phFontData, this.paragraph);
            if (phFontData.IsBold is not null)
            {
                phFontData.IsBold = new BooleanValue(value);
            }
            else
            {
                var aEndParaRPr = this.aText.Parent.NextSibling<DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties>();
                if (aEndParaRPr != null)
                {
                    aEndParaRPr.Bold = new BooleanValue(value);
                }
                else
                {
                    aRunPr = new DocumentFormat.OpenXml.Drawing.RunProperties { Bold = new BooleanValue(value) };
                    this.aText.Parent.InsertAt(aRunPr, 0); // append to <a:r>
                }
            }
        }
    }

    private void SetItalicFlag(bool isItalic)
    {
        var aTextParent = this.aText.Parent!;
        var aRunPr = aTextParent.GetFirstChild<DocumentFormat.OpenXml.Drawing.RunProperties>();
        if (aRunPr != null)
        {
            aRunPr.Italic = new BooleanValue(isItalic);
        }
        else
        {
            var aEndParaRPr = aTextParent.NextSibling<DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties>();
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

    private void SetFontSize(int newFontSize)
    {
        var parent = this.aText.Parent!;
        var aRunPr = parent.GetFirstChild<DocumentFormat.OpenXml.Drawing.RunProperties>();
        if (aRunPr == null)
        {
            var builder = new ARunPropertiesBuilder();
            aRunPr = builder.Build();
            parent.InsertAt(aRunPr, 0);
        }

        aRunPr.FontSize = newFontSize * 100;
    }
}
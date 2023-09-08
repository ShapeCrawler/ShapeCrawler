using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Services;
using ShapeCrawler.Shared;
using ShapeCrawler.Wrappers;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Texts;

internal sealed record TextFrame : ITextFrame
{
    private readonly SlidePart sdkSlidePart;
    private readonly TypedOpenXmlCompositeElement textBody;
    private readonly ResetableLazy<string> text;
    private readonly ResetableLazy<Paragraphs> paragraphs;
    
    internal TextFrame(SlidePart sdkSlidePart,TypedOpenXmlCompositeElement textBody)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.textBody = textBody;
        this.text = new ResetableLazy<string>(this.GetText);
        this.paragraphs = new ResetableLazy<Paragraphs>(() => new Paragraphs(sdkSlidePart, this.textBody.Elements<A.Paragraph>()));
    }

    internal event Action? TextChanged;

    public IParagraphCollection Paragraphs => this.paragraphs.Value;

    public string Text
    {
        get => this.text.Value;
        set => this.UpdateText(value);
    }

    public SCAutofitType AutofitType
    {
        get => this.GetAutofitType();
        set => this.SetAutofitType(value);
    }

    public double LeftMargin
    {
        get => this.GetLeftMargin();
        set => this.SetLeftMargin(value);
    }

    public double RightMargin
    {
        get => this.GetRightMargin();
        set => this.SetRightMargin(value);
    }

    public double TopMargin
    {
        get => this.GetTopMargin();
        set => this.SetTopMargin(value);
    }

    public double BottomMargin
    {
        get => this.GetBottomMargin();
        set => this.SetBottomMargin(value);
    }

    public bool TextWrapped => this.IsTextWrapped();

    internal void OnParagraphTextChanged()
    {
        this.text.Reset();
        this.TextChanged?.Invoke();
    }

    internal void Draw(SKCanvas slideCanvas, float shapeX, float shapeY)
    {
        using var paint = new SKPaint();
        paint.Color = SKColors.Black;
        var firstPortion = this.paragraphs.Value.First().Portions.First();
        paint.TextSize = firstPortion.Font!.Size;
        var typeFace = SKTypeface.FromFamilyName(firstPortion.Font.LatinName); 
        paint.Typeface = typeFace;
        float leftMarginPx = UnitConverter.CentimeterToPixel(this.LeftMargin);
        float topMarginPx = UnitConverter.CentimeterToPixel(this.TopMargin);
        float fontHeightPx = UnitConverter.PointToPixel(16);
        float x = shapeX + leftMarginPx; 
        float y = shapeY + topMarginPx + fontHeightPx; 
        slideCanvas.DrawText(this.Text, x, y, paint);
    }

    private void SetAutofitType(SCAutofitType newType)
    {
        var currentType = this.AutofitType;
        if (currentType == newType)
        {
            return;
        }

        var aBodyPr = this.textBody.GetFirstChild<A.BodyProperties>() !;
        var dontAutofit = aBodyPr.GetFirstChild<A.NoAutoFit>();
        var shrink = aBodyPr.GetFirstChild<A.NormalAutoFit>();
        var resize = aBodyPr.GetFirstChild<A.ShapeAutoFit>();

        switch (newType)
        {
            case SCAutofitType.None:
                shrink?.Remove();
                resize?.Remove();
                dontAutofit = new A.NoAutoFit();
                aBodyPr.Append(dontAutofit);
                break;
            case SCAutofitType.Shrink:
                dontAutofit?.Remove();
                resize?.Remove();
                shrink = new A.NormalAutoFit();
                aBodyPr.Append(shrink);
                break;
            case SCAutofitType.Resize:
            {
                dontAutofit?.Remove();
                shrink?.Remove();
                resize = new A.ShapeAutoFit();
                aBodyPr.Append(resize);
                break;
            }

            default:
                throw new ArgumentOutOfRangeException(nameof(newType), newType, null);
        }
    }

    private double GetLeftMargin()
    {
        var bodyProperties = this.textBody.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.LeftInset;
        return ins is null ? SCConstants.DefaultLeftAndRightMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }
    
    private double GetRightMargin()
    {
        var bodyProperties = this.textBody.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.RightInset;
        return ins is null ? SCConstants.DefaultLeftAndRightMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }

    private void SetLeftMargin(double centimetre)
    {
        var bodyProperties = this.textBody.GetFirstChild<A.BodyProperties>() !;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        bodyProperties.LeftInset = new Int32Value((int)emu);
    }

    private void SetRightMargin(double centimetre)
    {
        var bodyProperties = this.textBody.GetFirstChild<A.BodyProperties>() !;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        bodyProperties.RightInset = new Int32Value((int)emu);
    }
    
    private void SetTopMargin(double centimetre)
    {
        var bodyProperties = this.textBody.GetFirstChild<A.BodyProperties>() !;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        bodyProperties.TopInset = new Int32Value((int)emu);
    }

    private void SetBottomMargin(double centimetre)
    {
        var bodyProperties = this.textBody.GetFirstChild<A.BodyProperties>() !;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        bodyProperties.BottomInset = new Int32Value((int)emu);
    }

    private bool IsTextWrapped()
    {
        var aBodyPr = this.textBody.GetFirstChild<A.BodyProperties>() !;
        var wrap = aBodyPr.GetAttributes().FirstOrDefault(a => a.LocalName == "wrap");

        if (wrap == null)
        {
            return false;
        }

        if (wrap.Value == "none")
        {
            return false;
        }

        return true;
    }

    private double GetTopMargin()
    {
        var bodyProperties = this.textBody.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.TopInset;
        return ins is null ? SCConstants.DefaultTopAndBottomMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }

    private double GetBottomMargin()
    {
        var bodyProperties = this.textBody.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.BottomInset;
        return ins is null ? SCConstants.DefaultTopAndBottomMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }

    private SCAutofitType GetAutofitType()
    {
        var aBodyPr = this.textBody.GetFirstChild<A.BodyProperties>();
        if (aBodyPr!.GetFirstChild<A.NormalAutoFit>() != null)
        {
            return SCAutofitType.Shrink;
        }

        if (aBodyPr.GetFirstChild<A.ShapeAutoFit>() != null)
        {
            return SCAutofitType.Resize;
        }

        return SCAutofitType.None;
    }

    private void UpdateText(string newText)
    {
        var baseParagraph = this.Paragraphs.FirstOrDefault(p => p.Portions.Any());
        if (baseParagraph == null)
        {
            baseParagraph = this.Paragraphs.First();
            baseParagraph.Portions.AddText(newText);
        }

        var removingParagraphs = this.Paragraphs.Where(p => p != baseParagraph);
        this.Paragraphs.Remove(removingParagraphs);

        baseParagraph.Text = newText;

        if (this.AutofitType == SCAutofitType.Shrink)
        {
            this.ShrinkText(newText, baseParagraph);
        }
    }

    private void ShrinkText(string newText, IParagraph baseParagraph)
    {
        var popularPortion = baseParagraph.Portions.GroupBy(p => p.Font!.Size).OrderByDescending(x => x.Count())
            .First().First();
        var font = popularPortion.Font;

        var sdkPresDocument = (PresentationDocument)this.sdkSlidePart.OpenXmlPackage;
        var slideSize = new SlideSize(sdkPresDocument.PresentationPart!.Presentation.SlideSize!);
        var fontSize = FontService.GetAdjustedFontSize(newText, font!, slideSize.Width(), slideSize.Height());

        var paragraphInternal = (SlideParagraph)baseParagraph;
        paragraphInternal.SetFontSize(fontSize);
    }

    private string GetText()
    {
        var sb = new StringBuilder();
        sb.Append(this.Paragraphs[0].Text);

        var paragraphsCount = this.Paragraphs.Count;
        var index = 1; // we've already added the text of first paragraph
        while (index < paragraphsCount)
        {
            sb.AppendLine();
            sb.Append(this.Paragraphs[index].Text);

            index++;
        }

        return sb.ToString();
    }
}
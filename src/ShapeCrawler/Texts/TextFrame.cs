using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Positions;
using ShapeCrawler.Services;
using ShapeCrawler.ShapeCollection;
using ShapeCrawler.Shared;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Texts;

internal sealed class TextFrame : ITextFrame
{
    private readonly OpenXmlPart sdkTypedOpenXmlPart;
    private readonly OpenXmlElement sdkTextBody;

    internal TextFrame(OpenXmlPart sdkTypedOpenXmlPart, OpenXmlElement sdkTextBody)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.sdkTextBody = sdkTextBody;
    }

    public IParagraphs Paragraphs => new Paragraphs(this.sdkTypedOpenXmlPart, this.sdkTextBody);

    public string Text
    {
        get
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
        
        set
        {
            var paragraphs = this.Paragraphs.ToList();
            var paragraphWithPortion = paragraphs.FirstOrDefault(p => p.Portions.Any());
            if (paragraphWithPortion == null)
            {
                paragraphWithPortion = paragraphs.First();
                paragraphWithPortion.Portions.AddText(value);
            }
            else
            {
                var removingParagraphs = paragraphs.Where(p => p != paragraphWithPortion);
                foreach (var removingParagraph in removingParagraphs)
                {
                    removingParagraph.Remove();
                }

                paragraphWithPortion.Text = value;
            }

            if (this.AutofitType == AutofitType.Shrink)
            {
                this.ShrinkText(value, paragraphWithPortion);
            }

            this.ResizeParentShape();
        }
    }

    public AutofitType AutofitType
    {
        get
        {
            var aBodyPr = this.sdkTextBody.GetFirstChild<A.BodyProperties>();

            if (aBodyPr!.GetFirstChild<A.NormalAutoFit>() != null)
            {
                return AutofitType.Shrink;
            }

            if (aBodyPr.GetFirstChild<A.ShapeAutoFit>() != null)
            {
                return AutofitType.Resize;
            }

            return AutofitType.None;
        }
        
        set
        {
            var currentType = this.AutofitType;
            if (currentType == value)
            {
                return;
            }
            
            var aBodyPr = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
            var dontAutofit = aBodyPr.GetFirstChild<A.NoAutoFit>();
            var shrink = aBodyPr.GetFirstChild<A.NormalAutoFit>();
            var resize = aBodyPr.GetFirstChild<A.ShapeAutoFit>();
            
            switch (value)
            {
                case AutofitType.None:
                    shrink?.Remove();
                    resize?.Remove();
                    dontAutofit = new A.NoAutoFit();
                    aBodyPr.Append(dontAutofit);
                    break;
                case AutofitType.Shrink:
                    dontAutofit?.Remove();
                    resize?.Remove();
                    shrink = new A.NormalAutoFit();
                    aBodyPr.Append(shrink);
                    break;
                case AutofitType.Resize:
                {
                    dontAutofit?.Remove();
                    shrink?.Remove();
                    resize = new A.ShapeAutoFit();
                    aBodyPr.Append(resize);
                    this.ResizeParentShape();
                    break;
                }
            
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
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
    
    public string SDKXPath => new XmlPath(this.sdkTextBody).XPath;
    
    public void ResizeParentShape()
    {
        if (this.AutofitType != AutofitType.Resize)
        {
            return;
        }

        var baseParagraph = this.Paragraphs.First();
        var popularPortion = baseParagraph.Portions.OfType<TextParagraphPortion>().GroupBy(p => p.Font.Size)
            .OrderByDescending(x => x.Count())
            .First().First();
        var font = popularPortion.Font;

        var paint = new SKPaint();
        var fontSize = font.Size;
        paint.TextSize = fontSize;
        paint.Typeface = SKTypeface.FromFamilyName(font.LatinName);
        paint.IsAntialias = true;

        var lMarginPixel = UnitConverter.CentimeterToPixel(this.LeftMargin);
        var rMarginPixel = UnitConverter.CentimeterToPixel(this.RightMargin);
        var tMarginPixel = UnitConverter.CentimeterToPixel(this.TopMargin);
        var bMarginPixel = UnitConverter.CentimeterToPixel(this.BottomMargin);

        var textRect = default(SKRect);
        var text = this.Text;
        paint.MeasureText(text, ref textRect);
        var textWidth = textRect.Width;
        var textHeight = paint.TextSize;
        var shapeSize = new ShapeSize(this.sdkTypedOpenXmlPart, this.sdkTextBody.Ancestors<P.Shape>().First());
        var currentBlockWidth = shapeSize.Width() - lMarginPixel - rMarginPixel;
        var currentBlockHeight = shapeSize.Height() - tMarginPixel - bMarginPixel;

        this.UpdateShapeHeight(textWidth, currentBlockWidth, textHeight, tMarginPixel, bMarginPixel, currentBlockHeight, this.sdkTextBody.Parent!);
        this.UpdateShapeWidthIfNeeded(paint, lMarginPixel, rMarginPixel, this, this.sdkTextBody.Parent!);
    }

    internal void Draw(SKCanvas slideCanvas, float shapeX, float shapeY)
    {
        using var paint = new SKPaint();
        paint.Color = SKColors.Black;
        var firstPortion = this.Paragraphs.First().Portions.First();
        paint.TextSize = firstPortion.Font.Size;
        var typeFace = SKTypeface.FromFamilyName(firstPortion.Font.LatinName);
        paint.Typeface = typeFace;
        float leftMarginPx = UnitConverter.CentimeterToPixel(this.LeftMargin);
        float topMarginPx = UnitConverter.CentimeterToPixel(this.TopMargin);
        float fontHeightPx = UnitConverter.PointToPixel(16);
        float x = shapeX + leftMarginPx;
        float y = shapeY + topMarginPx + fontHeightPx;
        slideCanvas.DrawText(this.Text, x, y, paint);
    }

    private double GetLeftMargin()
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.LeftInset;
        return ins is null ? Constants.DefaultLeftAndRightMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }

    private double GetRightMargin()
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.RightInset;
        return ins is null ? Constants.DefaultLeftAndRightMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }

    private void SetLeftMargin(double centimetre)
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        bodyProperties.LeftInset = new Int32Value((int)emu);
    }

    private void SetRightMargin(double centimetre)
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        bodyProperties.RightInset = new Int32Value((int)emu);
    }

    private void SetTopMargin(double centimetre)
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        bodyProperties.TopInset = new Int32Value((int)emu);
    }

    private void SetBottomMargin(double centimetre)
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        bodyProperties.BottomInset = new Int32Value((int)emu);
    }

    private bool IsTextWrapped()
    {
        var aBodyPr = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var wrap = aBodyPr.GetAttributes().FirstOrDefault(a => a.LocalName == "wrap");

        if (wrap.Value == "none")
        {
            return false;
        }

        return true;
    }

    private double GetTopMargin()
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.TopInset;
        return ins is null ? Constants.DefaultTopAndBottomMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }

    private double GetBottomMargin()
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.BottomInset;
        return ins is null ? Constants.DefaultTopAndBottomMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }

    private void ShrinkText(string newText, IParagraph baseParagraph)
    {
        var popularPortion = baseParagraph.Portions.GroupBy(p => p.Font!.Size).OrderByDescending(x => x.Count())
            .First().First();
        var font = popularPortion.Font;

        var parent = this.sdkTextBody.Parent!;
        var shapeSize = new ShapeSize(this.sdkTypedOpenXmlPart, parent);
        var fontSize = FontService.GetAdjustedFontSize(newText, font, shapeSize.Width(), shapeSize.Height());

        var paragraphInternal = (Paragraph)baseParagraph;
        paragraphInternal.SetFontSize(fontSize);
    }

    private void UpdateShapeWidthIfNeeded(
        SKPaint paint, 
        int lMarginPixel, 
        int rMarginPixel, 
        TextFrame textFrame,
        OpenXmlElement parent)
    {
        if (!textFrame.TextWrapped)
        {
            var longerText = textFrame.Paragraphs
                .Select(x => new { x.Text, x.Text.Length })
                .OrderByDescending(x => x.Length)
                .First().Text;
            var paraTextRect = default(SKRect);
            var widthInPixels = paint.MeasureText(longerText, ref paraTextRect);
            
            // SkiaSharp uses 72 Dpi (https://stackoverflow.com/a/69916569/2948684), ShapeCrawler uses 96 Dpi.
            // 96/72=1.4
            const double Scale = 1.4;
            var newWidth = (int)(widthInPixels * Scale) + lMarginPixel + rMarginPixel;
            new ShapeSize(this.sdkTypedOpenXmlPart, parent).UpdateWidth(newWidth);
        }
    }

    private void UpdateShapeHeight(
        float textWidth,
        int currentBlockWidth,
        float textHeight,
        int tMarginPixel,
        int bMarginPixel,
        int currentBlockHeight,
        OpenXmlElement parent)
    {
        var requiredRowsCount = textWidth / currentBlockWidth;
        var integerPart = (int)requiredRowsCount;
        var fractionalPart = requiredRowsCount - integerPart;
        if (fractionalPart > 0)
        {
            integerPart++;
        }

        var requiredHeight = (integerPart * textHeight) + tMarginPixel + bMarginPixel;
        var newHeight = (int)requiredHeight + tMarginPixel + bMarginPixel + tMarginPixel + bMarginPixel;
        var position = new Position(this.sdkTypedOpenXmlPart, parent);
        var size = new ShapeSize(this.sdkTypedOpenXmlPart, parent);
        size.UpdateHeight(newHeight);

        // We should raise the shape up by the amount which is half of the increased offset.
        // PowerPoint does the same thing.
        var yOffset = (requiredHeight - currentBlockHeight) / 2;
        position.UpdateY((int)(position.Y() - yOffset));
    }
}
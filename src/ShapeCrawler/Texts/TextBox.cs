using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Positions;
using ShapeCrawler.ShapeCollection;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Texts;

internal sealed record TextBox : ITextBox
{
    private readonly OpenXmlPart sdkTypedOpenXmlPart;
    private readonly OpenXmlElement sdkTextBody;

    private TextVerticalAlignment? valignment;

    internal TextBox(OpenXmlPart sdkTypedOpenXmlPart, OpenXmlElement sdkTextBody)
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

        set => this.SetText(value);
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

    public decimal LeftMargin
    {
        get => this.GetLeftMargin();
        set => this.SetLeftMargin(value);
    }

    public decimal RightMargin
    {
        get => this.GetRightMargin();
        set => this.SetRightMargin(value);
    }

    public decimal TopMargin
    {
        get => this.GetTopMargin();
        set => this.SetTopMargin(value);
    }

    public decimal BottomMargin
    {
        get => this.GetBottomMargin();
        set => this.SetBottomMargin(value);
    }

    public bool TextWrapped => this.IsTextWrapped();

    public string SdkXPath => new XmlPath(this.sdkTextBody).XPath;

    public TextVerticalAlignment VerticalAlignment
    {
        get
        {
            if (this.valignment.HasValue)
            {
                return this.valignment.Value;
            }

            var aBodyPr = this.sdkTextBody.GetFirstChild<A.BodyProperties>();

            if (aBodyPr!.Anchor?.Value == A.TextAnchoringTypeValues.Center)
            {
                this.valignment = TextVerticalAlignment.Middle;
            }
            else if (aBodyPr.Anchor?.Value == A.TextAnchoringTypeValues.Bottom)
            {
                this.valignment = TextVerticalAlignment.Bottom;
            }
            else
            {
                this.valignment = TextVerticalAlignment.Top;
            }

            return this.valignment.Value;
        }

        set => this.SetVerticalAlignment(value);
    }

    private void SetVerticalAlignment(TextVerticalAlignment alignmentValue)
    {
        var aTextAlignmentTypeValue = alignmentValue switch
        {
            TextVerticalAlignment.Top => A.TextAnchoringTypeValues.Top,
            TextVerticalAlignment.Middle => A.TextAnchoringTypeValues.Center,
            TextVerticalAlignment.Bottom => A.TextAnchoringTypeValues.Bottom,
            _ => throw new ArgumentOutOfRangeException(nameof(alignmentValue))
        };

        var aBodyPr = this.sdkTextBody.GetFirstChild<A.BodyProperties>();

        if (aBodyPr is not null)
        {
            aBodyPr.Anchor = aTextAlignmentTypeValue;
        }

        this.valignment = alignmentValue;
    }
    
    private void SetText(string value)
    {
        var paragraphs = this.Paragraphs.ToList();
        var portionPara = paragraphs.FirstOrDefault(p => p.Portions.Any());
        if (portionPara == null)
        {
            portionPara = paragraphs.First();
            portionPara.Portions.AddText(value);
        }
        else
        {
            var removingParagraphs = paragraphs.Where(p => p != portionPara);
            foreach (var removingParagraph in removingParagraphs)
            {
                removingParagraph.Remove();
            }

            portionPara.Text = value;
        }

        if (this.AutofitType == AutofitType.Shrink)
        {
            this.ShrinkText(value, portionPara);
        }

        this.ResizeParentShape();
    }
    
    public void ResizeParentShape()
    {
        if (this.AutofitType != AutofitType.Resize)
        {
            return;
        }

        var lMarginPixel = UnitConverter.CentimeterToPixel(this.LeftMargin);
        var rMarginPixel = UnitConverter.CentimeterToPixel(this.RightMargin);
        var tMarginPixel = UnitConverter.CentimeterToPixel(this.TopMargin);
        var bMarginPixel = UnitConverter.CentimeterToPixel(this.BottomMargin);

        var shapeSize = new ShapeSize(this.sdkTypedOpenXmlPart, this.sdkTextBody.Ancestors<P.Shape>().First());
        var currentBlockWidth = shapeSize.Width() - lMarginPixel - rMarginPixel;
        var currentBlockHeight = shapeSize.Height() - tMarginPixel - bMarginPixel;

        decimal requiredHeight = 0;
        foreach (var paragraph in this.Paragraphs)
        {
            var paragraphPortion = paragraph.Portions.OfType<TextParagraphPortion>();
            if (!paragraphPortion.Any())
            {
                continue;
            }

            var popularPortion = paragraphPortion.GroupBy(p => p.Font.Size)
                .OrderByDescending(x => x.Count())
                .First().First();
            var scFont = popularPortion.Font;

            var text = paragraph.Text.ToUpper();

            var textWidth = new Text(text, scFont).PxWidth;
            var textHeight = scFont.Size;

            var requiredRowsCount = textWidth / currentBlockWidth;
            var integerPart = (int)requiredRowsCount;
            var fractionalPart = requiredRowsCount - integerPart;
            if (fractionalPart > 0)
            {
                integerPart++;
            }

            requiredHeight += integerPart * textHeight;

            // TODO
            // requiredHeight += (integerPart * textHeight) + (decimal)SpacingBefore + (decimal) SpacingAfter;
        }

        this.UpdateShapeHeight(requiredHeight, tMarginPixel, bMarginPixel, currentBlockHeight, this.sdkTextBody.Parent!);
        this.UpdateShapeWidthIfNeeded(lMarginPixel, rMarginPixel, this, this.sdkTextBody.Parent!);
    }

    private decimal GetLeftMargin()
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.LeftInset;
        return ins is null ? Constants.DefaultLeftAndRightMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }

    private decimal GetRightMargin()
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.RightInset;
        return ins is null ? Constants.DefaultLeftAndRightMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }

    private void SetLeftMargin(decimal centimetre)
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        bodyProperties.LeftInset = new Int32Value((int)emu);
    }

    private void SetRightMargin(decimal centimetre)
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        bodyProperties.RightInset = new Int32Value((int)emu);
    }

    private void SetTopMargin(decimal centimetre)
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        bodyProperties.TopInset = new Int32Value((int)emu);
    }

    private void SetBottomMargin(decimal centimetre)
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

    private decimal GetTopMargin()
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.TopInset;
        return ins is null ? Constants.DefaultTopAndBottomMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }

    private decimal GetBottomMargin()
    {
        var bodyProperties = this.sdkTextBody.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.BottomInset;
        return ins is null ? Constants.DefaultTopAndBottomMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }

    private void ShrinkText(string newText, IParagraph baseParagraph)
    {
        var popularFont = baseParagraph.Portions.GroupBy(paraPortion => paraPortion.Font!.Size).OrderByDescending(x => x.Count())
            .First().First().Font!;
        var shapeSize = new ShapeSize(this.sdkTypedOpenXmlPart, this.sdkTextBody.Parent!);
        
        var text = new Text(newText, popularFont);
        text.FitInto(shapeSize.Width(), shapeSize.Height());
        
        var internalPara = (Paragraph)baseParagraph;
        internalPara.SetFontSize((int)text.FontSize);
    }

    private void UpdateShapeWidthIfNeeded(
        decimal lMarginPixel,
        decimal rMarginPixel,
        TextBox textBox,
        OpenXmlElement parent)
    {
        if (!textBox.TextWrapped)
        {
            var longerText = textBox.Paragraphs
                .Select(x => new { x.Text, x.Text.Length })
                .OrderByDescending(x => x.Length)
                .First().Text;

            var baseParagraph = this.Paragraphs.First();
            var popularPortion = baseParagraph.Portions.OfType<TextParagraphPortion>().GroupBy(p => p.Font.Size)
                .OrderByDescending(x => x.Count())
                .First().First();
            var font = popularPortion.Font;

            var widthInPixels = new Text(longerText, font).PxWidth;
            
            var newWidth = (int)(widthInPixels * 1.4M) // SkiaSharp uses 72 Dpi (https://stackoverflow.com/a/69916569/2948684), ShapeCrawler uses 96 Dpi. 96/72 = 1.4 
                           + lMarginPixel + rMarginPixel;
            new ShapeSize(this.sdkTypedOpenXmlPart, parent).UpdateWidth(newWidth);
        }
    }
    
    private void UpdateShapeHeight(
        decimal textHeight,
        decimal tMarginPixel,
        decimal bMarginPixel,
        decimal currentBlockHeight,
        OpenXmlElement parent)
    {
        var requiredHeight = textHeight + tMarginPixel + bMarginPixel;
        var newHeight = requiredHeight + tMarginPixel + bMarginPixel + tMarginPixel + bMarginPixel;
        var position = new Position(this.sdkTypedOpenXmlPart, parent);
        var size = new ShapeSize(this.sdkTypedOpenXmlPart, parent);
        size.UpdateHeight(newHeight);

        // We should raise the shape up by the amount which is half of the increased offset.
        // PowerPoint does the same thing.
        var yOffset = (requiredHeight - currentBlockHeight) / 2;
        position.UpdateY(position.Y() - yOffset);
    }
}
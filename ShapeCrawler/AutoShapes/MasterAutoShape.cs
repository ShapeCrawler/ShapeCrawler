using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Factories;
using ShapeCrawler.Services;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.AutoShapes;

/// <summary>
///     Represents an AutoShape on a Slide Master.
/// </summary>
internal class MasterAutoShape : MasterShape, IAutoShape, ITextFrameContainer, IFontDataReader
{
    private readonly ResettableLazy<Dictionary<int, FontData>> lvlToFontData;
    private readonly Lazy<ShapeFill> shapeFill;
    private readonly Lazy<TextFrame?> textBox;
    private readonly P.Shape pShape;

    internal MasterAutoShape(SCSlideMaster slideMasterInternal, P.Shape pShape)
        : base(pShape, slideMasterInternal)
    {
        this.textBox = new Lazy<TextFrame?>(this.GetTextBox);
        this.shapeFill = new Lazy<ShapeFill>(this.TryGetFill);
        this.lvlToFontData = new ResettableLazy<Dictionary<int, FontData>>(this.GetLvlToFontData);
        this.pShape = pShape;
    }

    #region Public Properties

    public Shape Shape => this;

    public ITextFrame? TextFrame => this.textBox.Value;

    public IShapeFill Fill => this.shapeFill.Value;

    public override SCShapeType ShapeType => SCShapeType.AutoShape;

    #endregion Public Properties

    private Dictionary<int, FontData> LvlToFontData => this.lvlToFontData.Value;

    public void FillFontData(int paragraphLvl, ref FontData fontData)
    {
        if (this.LvlToFontData.TryGetValue(paragraphLvl, out FontData masterFontData) && !fontData.IsFilled())
        {
            masterFontData.Fill(fontData);
            return;
        }

        var pTextStyles = this.SlideMasterInternal.PSlideMaster.TextStyles!;
        if (this.Placeholder!.Type != SCPlaceholderType.Title)
        {
            return;
        }

        var titleFontSize = pTextStyles.TitleStyle!.Level1ParagraphProperties!
            .GetFirstChild<A.DefaultRunProperties>()!.FontSize!.Value;
        if (fontData.FontSize is null)
        {
            fontData.FontSize = new Int32Value(titleFontSize);
        }
    }

    private Dictionary<int, FontData> GetLvlToFontData() // TODO: duplicate code in LayoutAutoShape
    {
        var texBody = this.pShape.TextBody!;
        var lvlToFontData = FontDataParser.FromCompositeElement(texBody.ListStyle!);

        if (!lvlToFontData.Any())
        {
            var endParaRunPrFs = texBody.GetFirstChild<A.Paragraph>()!
                .GetFirstChild<A.EndParagraphRunProperties>()?.FontSize;
            if (endParaRunPrFs is not null)
            {
                var fontData = new FontData
                {
                    FontSize = endParaRunPrFs
                };
                lvlToFontData.Add(1, fontData);
            }
        }

        return lvlToFontData;
    }

    private TextFrame? GetTextBox() // TODO: duplicate code in LayoutAutoShape
    {
        P.TextBody pTextBody = this.PShapeTreesChild.GetFirstChild<P.TextBody>();
        if (pTextBody == null)
        {
            return null;
        }

        IEnumerable<A.Text> aTexts = pTextBody.Descendants<A.Text>();
        if (aTexts.Sum(t => t.Text.Length) > 0)
        {
            return new TextFrame(this, pTextBody);
        }

        return null;
    }

    private ShapeFill TryGetFill() // TODO: duplicate code in LayoutAutoShape
    {
        throw new NotImplementedException();
    }
}
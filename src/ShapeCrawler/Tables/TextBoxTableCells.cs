using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal sealed record TextBoxTableCells : ITextBox
{
    private readonly A.TableCell tableCell;
    private readonly OpenXmlPart sdkTypedOpenXmlPart;

    private TextVerticalAlignment? valignment;

    internal TextBoxTableCells(OpenXmlPart sdkTypedOpenXmlPart, A.TableCell tableCell)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.tableCell = tableCell;
    }

    public TextVerticalAlignment VerticalAlignment
    {
        get
        {
            if (this.valignment.HasValue)
            {
                return this.valignment.Value;
            }

            var aBodyPr = this.tableCell.TableCellProperties!;

            aBodyPr!.Anchor ??= A.TextAnchoringTypeValues.Top;

            if (aBodyPr!.Anchor!.Value == A.TextAnchoringTypeValues.Center)
            {
                this.valignment = TextVerticalAlignment.Middle;
            }
            else if (aBodyPr!.Anchor!.Value == A.TextAnchoringTypeValues.Bottom)
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

        var aCellProperties = this.tableCell.TableCellProperties!;
        aCellProperties.Anchor = aTextAlignmentTypeValue;
        this.valignment = alignmentValue;
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

    public IParagraphs Paragraphs => new Paragraphs(this.sdkTypedOpenXmlPart, this.tableCell.TextBody!);

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

    private void SetText(string value)
    {
        var paragraphs = this.Paragraphs.ToList();
        var portionPara = paragraphs.FirstOrDefault(p => p.Portions.Count != 0);
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
    }

    public AutofitType AutofitType { get => AutofitType.None; set => throw new NotSupportedException(); }

    public bool TextWrapped => true;

    public string SdkXPath => new XmlPath(this.tableCell.TextBody!).XPath;

    private decimal GetLeftMargin()
    {
        var cellProperty = this.tableCell.TableCellProperties!;
        var margin = cellProperty.LeftMargin;
        return margin is null ? Constants.DefaultLeftAndRightMargin : UnitConverter.EmuToCentimeter(margin.Value);
    }

    private decimal GetRightMargin()
    {
        var cellProperty = this.tableCell.TableCellProperties!;
        var margin = cellProperty.RightMargin;
        return margin is null ? Constants.DefaultLeftAndRightMargin : UnitConverter.EmuToCentimeter(margin.Value);
    }

    private void SetLeftMargin(decimal centimetre)
    {
        var cellProperties = this.tableCell.TableCellProperties!;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        cellProperties.LeftMargin = new Int32Value((int)emu);
    }

    private void SetRightMargin(decimal centimetre)
    {
        var cellProperties = this.tableCell.TableCellProperties!;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        cellProperties.RightMargin = new Int32Value((int)emu);
    }

    private void SetTopMargin(decimal centimetre)
    {
        var cellProperties = this.tableCell.TableCellProperties!;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        cellProperties.TopMargin = new Int32Value((int)emu);
    }

    private void SetBottomMargin(decimal centimetre)
    {
        var cellProperties = this.tableCell.TableCellProperties!;
        var emu = UnitConverter.CentimeterToEmu(centimetre);
        cellProperties.BottomMargin = new Int32Value((int)emu);
    }

    private decimal GetTopMargin()
    {
        var cellProperties = this.tableCell.TableCellProperties!;
        var margins = cellProperties.TopMargin;
        return margins is null ? Constants.DefaultTopAndBottomMargin : UnitConverter.EmuToCentimeter(margins.Value);
    }

    private decimal GetBottomMargin()
    {
        var cellProperties = this.tableCell.TableCellProperties!;
        var margins = cellProperties.BottomMargin;
        return margins is null ? Constants.DefaultTopAndBottomMargin : UnitConverter.EmuToCentimeter(margins.Value);
    }
}
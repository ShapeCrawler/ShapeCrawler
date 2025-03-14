using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Texts;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal sealed record TextBoxTableCells : ITextBox
{
    private readonly A.TableCell tableCell;
    private readonly OpenXmlPart sdkTypedOpenXmlPart;

    private TextVerticalAlignment? vAlignment;

    internal TextBoxTableCells(OpenXmlPart sdkTypedOpenXmlPart, A.TableCell tableCell)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.tableCell = tableCell;
    }

    public TextVerticalAlignment VerticalAlignment
    {
        get
        {
            if (this.vAlignment.HasValue)
            {
                return this.vAlignment.Value;
            }

            var aBodyPr = this.tableCell.TableCellProperties!;

            aBodyPr!.Anchor ??= A.TextAnchoringTypeValues.Top;

            if (aBodyPr!.Anchor!.Value == A.TextAnchoringTypeValues.Center)
            {
                this.vAlignment = TextVerticalAlignment.Middle;
            }
            else if (aBodyPr!.Anchor!.Value == A.TextAnchoringTypeValues.Bottom)
            {
                this.vAlignment = TextVerticalAlignment.Bottom;
            }
            else
            {
                this.vAlignment = TextVerticalAlignment.Top;
            }

            return this.vAlignment.Value;
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
        this.vAlignment = alignmentValue;
    }

    public float LeftMargin
    {
        get => new LeftRightMargin(this.tableCell.TableCellProperties!.LeftMargin).Value;
        set => this.SetLeftMargin(value);
    }

    public float RightMargin
    {
        get => new LeftRightMargin(this.tableCell.TableCellProperties!.RightMargin).Value;
        set => this.SetRightMargin(value);
    }

    public float TopMargin
    {
        get => new TopBottomMargin(this.tableCell.TableCellProperties!.TopMargin).Value;
        set => this.SetTopMargin(value);
    }

    public float BottomMargin
    {
        get => new TopBottomMargin(this.tableCell.TableCellProperties!.BottomMargin).Value;
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

    private void SetLeftMargin(float points)
    {
        var cellProperties = this.tableCell.TableCellProperties!;
        var emu = new Points(points).AsEmus();
        cellProperties.LeftMargin = new Int32Value((int)emu);
    }

    private void SetRightMargin(float points)
    {
        var cellProperties = this.tableCell.TableCellProperties!;
        var emu = new Points(points).AsEmus();
        cellProperties.RightMargin = new Int32Value((int)emu);
    }

    private void SetTopMargin(float points)
    {
        var cellProperties = this.tableCell.TableCellProperties!;
        var emu = new Points(points).AsEmus();
        cellProperties.TopMargin = new Int32Value((int)emu);
    }

    private void SetBottomMargin(float points)
    {
        var cellProperties = this.tableCell.TableCellProperties!;
        var emu = new Points(points).AsEmus();
        cellProperties.BottomMargin = new Int32Value((int)emu);
    }
}
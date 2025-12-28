using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using ShapeCrawler.Paragraphs;
using ShapeCrawler.Texts;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tables;

internal sealed class TableCellTextBox(A.TableCell aTableCell) : ITextBox
{
    private TextVerticalAlignment? vAlignment;
    private TextDirection? textDirection;

    public TextVerticalAlignment VerticalAlignment
    {
        get
        {
            if (this.vAlignment.HasValue)
            {
                return this.vAlignment.Value;
            }

            var aBodyPr = aTableCell.TableCellProperties!;
            aBodyPr.Anchor ??= A.TextAnchoringTypeValues.Top;

            if (aBodyPr.Anchor!.Value == A.TextAnchoringTypeValues.Center)
            {
                this.vAlignment = TextVerticalAlignment.Middle;
            }
            else if (aBodyPr.Anchor!.Value == A.TextAnchoringTypeValues.Bottom)
            {
                this.vAlignment = TextVerticalAlignment.Bottom;
            }
            else
            {
                this.vAlignment = TextVerticalAlignment.Top;
            }

            return this.vAlignment.Value;
        }

        set
        {
            var aTextAlignmentTypeValue = value switch
            {
                TextVerticalAlignment.Top => A.TextAnchoringTypeValues.Top,
                TextVerticalAlignment.Middle => A.TextAnchoringTypeValues.Center,
                TextVerticalAlignment.Bottom => A.TextAnchoringTypeValues.Bottom,
                _ => throw new ArgumentOutOfRangeException(nameof(value))
            };

            var aCellProperties = aTableCell.TableCellProperties!;
            aCellProperties.Anchor = aTextAlignmentTypeValue;
            this.vAlignment = value;
        }
    }

    public decimal LeftMargin
    {
        get => new LeftRightMargin(aTableCell.TableCellProperties!.LeftMargin).Value;
        set
        {
            var cellProperties = aTableCell.TableCellProperties!;
            var emu = new Points(value).AsEmus();
            cellProperties.LeftMargin = new Int32Value((int)emu);
        }
    }

    public decimal RightMargin
    {
        get => new LeftRightMargin(aTableCell.TableCellProperties!.RightMargin).Value;
        set
        {
            var cellProperties = aTableCell.TableCellProperties!;
            var emu = new Points(value).AsEmus();
            cellProperties.RightMargin = new Int32Value((int)emu);
        }
    }

    public decimal TopMargin
    {
        get => new TopBottomMargin(aTableCell.TableCellProperties!.TopMargin).Value;
        set
        {
            var cellProperties = aTableCell.TableCellProperties!;
            var emu = new Points(value).AsEmus();
            cellProperties.TopMargin = new Int32Value((int)emu);
        }
    }

    public decimal BottomMargin
    {
        get => new TopBottomMargin(aTableCell.TableCellProperties!.BottomMargin).Value;
        set
        {
            var cellProperties = aTableCell.TableCellProperties!;
            var emu = new Points(value).AsEmus();
            cellProperties.BottomMargin = new Int32Value((int)emu);
        }
    }

    public IParagraphCollection Paragraphs => new ParagraphCollection(aTableCell.TextBody!);

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
    }

    public AutofitType AutofitType { get => AutofitType.None; set => throw new NotSupportedException(); }

    public bool TextWrapped => true;

    public string SdkXPath => new XmlPath(aTableCell.TextBody!).XPath;

    public TextDirection TextDirection
    {
        get
        {
            if (this.textDirection.HasValue)
            {
                return this.textDirection.Value;
            }

            var textPositionValue = aTableCell.TableCellProperties!.Vertical?.Value;

            if (textPositionValue == A.TextVerticalValues.Vertical)
            {
                this.textDirection = TextDirection.Rotate90;
            }
            else if (textPositionValue == A.TextVerticalValues.Vertical270)
            {
                this.textDirection = TextDirection.Rotate270;
            }
            else if (textPositionValue == A.TextVerticalValues.WordArtVertical)
            {
                this.textDirection = TextDirection.Stacked;
            }
            else
            {
                this.textDirection = TextDirection.Horizontal;
            }

            return this.textDirection.Value;
        }

        set => this.SetTextDirection(value);
    }

    public void SetMarkdownText(string text)
    {
        throw new NotImplementedException();
    }

    public void SetText(string text)
    {
        var textLines = SplitLines(text);

        var firstParagraph = this.EnsureFirstParagraph();
        this.RemoveExtraParagraphs();
        ClearParagraphPortions(firstParagraph);

        if (textLines.Length > 0)
        {
            firstParagraph.Portions.AddText(textLines[0]);
        }

        this.AddRemainingLinesAsParagraphs(textLines);

        this.AdjustRowHeightForCurrentContent();
    }

    private static string[] SplitLines(string text)
    {
        return text.Split([Environment.NewLine, "\n"], StringSplitOptions.None);
    }

    private static A.Table? GetATable(A.TableRow aTableRow)
    {
        var graphicFrame = aTableRow.Ancestors<P.GraphicFrame>().FirstOrDefault();
        return graphicFrame?.GetFirstChild<A.Graphic>()?.GraphicData?.GetFirstChild<A.Table>();
    }

    private static void ClearParagraphPortions(IParagraph paragraph)
    {
        foreach (var portion in paragraph.Portions.ToList())
        {
            portion.Remove();
        }
    }

    private IParagraph EnsureFirstParagraph()
    {
        var existingParagraphs = this.Paragraphs.ToList();
        var firstParagraph = existingParagraphs.FirstOrDefault();
        if (firstParagraph != null)
        {
            return firstParagraph;
        }

        this.Paragraphs.Add();
        return this.Paragraphs[0];
    }

    private void RemoveExtraParagraphs()
    {
        var existingParagraphs = this.Paragraphs.ToList();
        foreach (var paragraph in existingParagraphs.Skip(1))
        {
            paragraph.Remove();
        }
    }

    private void AddRemainingLinesAsParagraphs(string[] textLines)
    {
        for (var i = 1; i < textLines.Length; i++)
        {
            this.Paragraphs.Add();
            var newParagraph = this.Paragraphs[this.Paragraphs.Count - 1];
            ClearParagraphPortions(newParagraph);
            newParagraph.Portions.AddText(textLines[i]);
        }
    }

    private void AdjustRowHeightForCurrentContent()
    {
        var aTableRow = aTableCell.Ancestors<A.TableRow>().FirstOrDefault();
        if (aTableRow == null)
        {
            return;
        }

        var aTable = GetATable(aTableRow);
        if (aTable?.TableGrid == null)
        {
            return;
        }

        var colIndex = this.GetColumnIndex(aTableRow);
        if (colIndex < 0)
        {
            return;
        }

        var widthCapacity = this.GetWidthCapacityPoints(aTable, colIndex);
        if (widthCapacity <= 0)
        {
            return;
        }

        var textHeight = this.CalculateTextHeight(widthCapacity);
        var requiredHeight = textHeight + this.TopMargin + this.BottomMargin;
        var currentRowHeight = new Emus(aTableRow.Height!.Value).AsPoints();
        if (requiredHeight <= currentRowHeight)
        {
            return;
        }

        var rowIndex = aTable.Elements<A.TableRow>().ToList().IndexOf(aTableRow);
        var scRow = new ShapeCrawler.TableRow(aTableRow, rowIndex);
        scRow.SetHeight(requiredHeight);
    }

    private int GetColumnIndex(A.TableRow aTableRow)
    {
        var cellsInRow = aTableRow.Elements<A.TableCell>().ToList();
        return cellsInRow.IndexOf(aTableCell);
    }

    private decimal GetWidthCapacityPoints(A.Table aTable, int colIndex)
    {
        var gridColumns = aTable.TableGrid!.Elements<A.GridColumn>().ToList();
        if (colIndex >= gridColumns.Count)
        {
            return 0;
        }

        var columnWidthPts = new Emus(gridColumns[colIndex].Width!.Value).AsPoints();
        return columnWidthPts - this.LeftMargin - this.RightMargin;
    }

    private decimal CalculateTextHeight(decimal widthCapacity)
    {
        decimal textHeight = 0;
        foreach (var paragraph in this.Paragraphs)
        {
            var paragraphPortions = paragraph.Portions.OfType<TextParagraphPortion>();
            if (!paragraphPortions.Any())
            {
                continue;
            }

            var popularPortion = paragraphPortions
                .GroupBy(p => p.Font.Size)
                .OrderByDescending(g => g.Count())
                .First().First();
            var scFont = popularPortion.Font;

            var paragraphText = paragraph.Text.ToUpper();
            var paragraphTextWidth = new Text(paragraphText, scFont).Width;
            var requiredRowsCount = paragraphTextWidth / widthCapacity;
            var intRequiredRowsCount = (int)Math.Ceiling(requiredRowsCount);
            if (intRequiredRowsCount == 0 && paragraphTextWidth > 0)
            {
                intRequiredRowsCount = 1;
            }

            textHeight += intRequiredRowsCount * scFont.Size;
        }

        return textHeight;
    }

    private void SetTextDirection(TextDirection value)
    {
        aTableCell.TableCellProperties!.Vertical = value switch
        {
            TextDirection.Rotate90 => A.TextVerticalValues.Vertical,
            TextDirection.Rotate270 => A.TextVerticalValues.Vertical270,
            TextDirection.Stacked => A.TextVerticalValues.WordArtVertical,
            _ => A.TextVerticalValues.Horizontal
        };

        this.TextDirection = value;
    }
}
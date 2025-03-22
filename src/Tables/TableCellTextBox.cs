using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using ShapeCrawler.Texts;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

internal sealed class TableCellTextBox(A.TableCell aTableCell): ITextBox
{
    private TextVerticalAlignment? vAlignment;

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

        set => this.SetText(value);
    }

    public AutofitType AutofitType { get => AutofitType.None; set => throw new NotSupportedException(); }

    public bool TextWrapped => true;

    public string SdkXPath => new XmlPath(aTableCell.TextBody!).XPath;
    
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
}
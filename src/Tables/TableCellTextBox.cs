using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using ShapeCrawler.Texts;
using ShapeCrawler.Paragraphs;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tables;

internal sealed class TableCellTextBox(A.TableCell aTableCell): ITextBox
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

    public string SDKXPath => new XmlPath(aTableCell.TextBody!).XPath;

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
        // Split text by newlines to handle multiple paragraphs
        var textLines = text.Split([Environment.NewLine, "\n"], StringSplitOptions.None);
        
        // Clear all existing paragraphs
        var existingParagraphs = this.Paragraphs.ToList();
        var firstParagraph = existingParagraphs.FirstOrDefault();
        
        // Keep only the first paragraph and clear all its portions
        if (firstParagraph == null)
        {
            // Create a paragraph if none exists
            this.Paragraphs.Add();
            firstParagraph = this.Paragraphs[0];
        }
        else
        {
            // Remove all paragraphs after the first one
            foreach (var p in existingParagraphs.Skip(1))
            {
                p.Remove();
            }
            
            // Clear all portions in the first paragraph
            foreach (var portion in firstParagraph.Portions.ToList())
            {
                portion.Remove();
            }
        }
        
        // Add the first line of text to the first paragraph
        if (textLines.Length > 0)
        {
            firstParagraph.Portions.AddText(textLines[0]);
        }
        
        // Create a new paragraph for each additional line
        for (int i = 1; i < textLines.Length; i++)
        {
            // Add a new paragraph
            this.Paragraphs.Add();
            
            // Get the newly created paragraph
            var newParagraph = this.Paragraphs[i];
            
            // Clear any existing portions (since it was cloned from the previous paragraph)
            foreach (var portion in newParagraph.Portions.ToList())
            {
                portion.Remove();
            }
            
            // Add the text for this line
            newParagraph.Portions.AddText(textLines[i]);
        }

        // Adjust the row height if the text wraps to multiple lines
        var aTableRow = aTableCell.Ancestors<A.TableRow>().FirstOrDefault();
        if (aTableRow != null)
        {
            var graphicFrame = aTableRow.Ancestors<P.GraphicFrame>().FirstOrDefault();
            var aTable = graphicFrame?.GetFirstChild<A.Graphic>()?.GraphicData?.GetFirstChild<A.Table>();
            if (aTable?.TableGrid != null)
            {
                var cellsInRow = aTableRow.Elements<A.TableCell>().ToList();
                var colIndex = cellsInRow.IndexOf(aTableCell);
                if (colIndex >= 0)
                {
                    var gridColumns = aTable.TableGrid.Elements<A.GridColumn>().ToList();
                    if (colIndex < gridColumns.Count)
                    {
                        var columnWidthPts = new Emus(gridColumns[colIndex].Width!.Value).AsPoints();
                        var widthCapacity = columnWidthPts - this.LeftMargin - this.RightMargin;

                        if (widthCapacity > 0)
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
                                // SkiaSharp uses 72 DPI while ShapeCrawler uses 96 DPI; adjust width accordingly (96/72 ≈ 1.4)
                                var paragraphTextWidth = new Text(paragraphText, scFont).Width * 1.4M;
                                var requiredRowsCount = paragraphTextWidth / widthCapacity;
                                var intRequiredRowsCount = (int)Math.Ceiling(requiredRowsCount);
                                if (intRequiredRowsCount == 0 && paragraphTextWidth > 0)
                                {
                                    intRequiredRowsCount = 1;
                                }

                                textHeight += intRequiredRowsCount * (int)scFont.Size;
                            }

                            var requiredHeight = textHeight + this.TopMargin + this.BottomMargin;
                            var currentRowHeight = new Emus(aTableRow.Height!.Value).AsPoints();
                            if (requiredHeight > currentRowHeight)
                            {
                                var rowIndex = aTable.Elements<A.TableRow>().ToList().IndexOf(aTableRow);
                                var scRow = new ShapeCrawler.TableRow(aTableRow, rowIndex);
                                scRow.SetHeight(requiredHeight);
                            }
                        }
                    }
                }
            }
        }
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
using System.Collections.Generic;
using ShapeCrawler.Drawing;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tables;
using ShapeCrawler.Texts;
using ShapeCrawler.Units;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class TableShape : DrawingShape
{
    private const string MediumStyle2Accent1Guid = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

    internal TableShape(Position position, ShapeSize shapeSize, ShapeId shapeId, P.GraphicFrame pGraphicFrame)
        : base(position, shapeSize, shapeId, pGraphicFrame)
    {
        var aTable = pGraphicFrame.GetFirstChild<A.Graphic>()!.GetFirstChild<A.GraphicData>()!
            .GetFirstChild<A.Table>()!;
        this.Table = new Table(
            new TableRowCollection(pGraphicFrame),
            new TableColumnCollection(pGraphicFrame),
            new TableStyleOptions(aTable.TableProperties!),
            pGraphicFrame);
    }

    public override ITable? Table { get; }

    public override ShapeContentType ContentType => ShapeContentType.Table;

    public override decimal Width
    {
        get => base.Width;
        set
        {
            var percentNewWidth = value / this.Width;
            base.Width = value;

            foreach (var tableColumn in this.Table!.Columns)
            {
                tableColumn.Width *= percentNewWidth;
            }
        }
    }

    public override decimal Height
    {
        get => base.Height;
        set
        {
            var percentNewHeight = value / this.Height;
            base.Height = value;

            foreach (var tableRow in this.Table!.Rows)
            {
                var row = (TableRow)tableRow;
                row.SetHeight((int)(row.Height * percentNewHeight));
            }
        }
    }

    internal override void Render(SKCanvas canvas)
    {
        var table = (Table)this.Table!;
        var rowTopPoints = this.Y;

        for (var rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++)
        {
            var colLeftPoints = this.X;
            this.RenderRow(canvas, table, rowIdx, ref colLeftPoints, rowTopPoints);
            rowTopPoints += table.Rows[rowIdx].Height;
        }
    }

    private static SKRect CreateRectFromPoints(decimal x, decimal y, decimal width, decimal height)
    {
        return new SKRect(
            (float)new Points(x).AsPixels(),
            (float)new Points(y).AsPixels(),
            (float)new Points(x + width).AsPixels(),
            (float)new Points(y + height).AsPixels());
    }

    private static void RenderCellText(SKCanvas canvas, TableCell cell, decimal x, decimal y, decimal w, decimal h, string? styleFontColorHex)
    {
        var aTextBody = cell.ATableCell.TextBody!;
        
        if (styleFontColorHex == null)
        {
            RenderTextWithoutStyleColor(canvas, aTextBody, x, y, w, h);
            return;
        }

        var modifiedRunProperties = ApplyStyleFontColor(aTextBody, styleFontColorHex);
        
        try
        {
            RenderTextWithoutStyleColor(canvas, aTextBody, x, y, w, h);
        }
        finally
        {
            RestoreOriginalFontColors(modifiedRunProperties);
        }
    }

    private static void RenderTextWithoutStyleColor(SKCanvas canvas, A.TextBody aTextBody, decimal x, decimal y, decimal w, decimal h)
    {
        var textBoxMargins = new TextBoxMargins(aTextBody);
        var drawingTextBox = new DrawingTextBox(textBoxMargins, aTextBody);
        drawingTextBox.Render(canvas, x, y, w, h);
    }

    private static List<(A.RunProperties RunProp, A.SolidFill? OriginalFill)> ApplyStyleFontColor(A.TextBody aTextBody, string styleFontColorHex)
    {
        var modifiedRunProperties = new List<(A.RunProperties RunProp, A.SolidFill? OriginalFill)>();

        foreach (var aParagraph in aTextBody.Elements<A.Paragraph>())
        {
            foreach (var aRun in aParagraph.Elements<A.Run>())
            {
                var aRunPr = EnsureRunProperties(aRun);
                var originalFill = aRunPr.GetFirstChild<A.SolidFill>();
                modifiedRunProperties.Add((aRunPr, originalFill));

                ApplyColorToRun(aRunPr, originalFill, styleFontColorHex);
            }
        }

        return modifiedRunProperties;
    }

    private static A.RunProperties EnsureRunProperties(A.Run aRun)
    {
        var aRunPr = aRun.GetFirstChild<A.RunProperties>();
        if (aRunPr == null)
        {
            aRunPr = new A.RunProperties();
            aRun.InsertAt(aRunPr, 0);
        }

        return aRunPr;
    }

    private static void ApplyColorToRun(A.RunProperties aRunPr, A.SolidFill? originalFill, string colorHex)
    {
        originalFill?.Remove();
        var newFill = new A.SolidFill();
        newFill.Append(new A.RgbColorModelHex { Val = colorHex });
        aRunPr.InsertAt(newFill, 0);
    }

    private static void RestoreOriginalFontColors(List<(A.RunProperties RunProp, A.SolidFill? OriginalFill)> modifiedRunProperties)
    {
        foreach (var (runProp, originalFill) in modifiedRunProperties)
        {
            var tempFill = runProp.GetFirstChild<A.SolidFill>();
            tempFill?.Remove();
            
            if (originalFill != null)
            {
                runProp.InsertAt(originalFill, 0);
            }
        }
    }
    
    private static (decimal Width, decimal Height) CalculateCellDimensions(Table table, TableCell cell, int rowIdx, int colIdx)
    {
        int gridSpan = cell.ATableCell.GridSpan?.Value ?? 1;
        int rowSpan = cell.ATableCell.RowSpan?.Value ?? 1;

        decimal cellTotalWidth = 0;
        for (int k = 0; k < gridSpan; k++)
        {
            cellTotalWidth += table.Columns[colIdx + k].Width;
        }

        decimal cellTotalHeight = 0;
        for (int k = 0; k < rowSpan; k++)
        {
            cellTotalHeight += table.Rows[rowIdx + k].Height;
        }

        return (cellTotalWidth, cellTotalHeight);
    }

    private void RenderRow(SKCanvas canvas, Table table, int rowIdx, ref decimal colLeftPoints, decimal rowTopPoints)
    {
        var row = table.Rows[rowIdx];
        var columns = table.Columns;

        for (var colIdx = 0; colIdx < columns.Count; colIdx++)
        {
            var cell = (TableCell)row.Cells[colIdx];

            // Render only if it's the top-left cell of a merge (or single cell)
            if (cell.RowIndex == rowIdx && cell.ColumnIndex == colIdx)
            {
                var cellDimensions = CalculateCellDimensions(table, cell, rowIdx, colIdx);
                this.RenderCell(canvas, cell, colLeftPoints, rowTopPoints, cellDimensions.Width, cellDimensions.Height);
            }

            colLeftPoints += columns[colIdx].Width;
        }
    }

    private void RenderCell(SKCanvas canvas, TableCell cell, decimal x, decimal y, decimal w, decimal h)
    {
        // 1. Resolve Fill
        var fillColor = this.GetCellFillColor(cell);
        
        // 2. Render Fill
        if (fillColor != null)
        {
            var rect = CreateRectFromPoints(x, y, w, h);

            using var paint = new SKPaint();
            paint.Color = fillColor.Value;
            paint.Style = SKPaintStyle.Fill;
            paint.IsAntialias = true;
            canvas.DrawRect(rect, paint);
        }

        // 3. Render Borders
        this.RenderBorders(canvas, x, y, w, h);

        // 4. Render Text with style font color
        if (cell.ATableCell.TextBody == null)
        {
            return;
        }

        var styleFontColorHex = this.GetStyleFontColorHex(cell);
        RenderCellText(canvas, cell, x, y, w, h, styleFontColorHex);
    }

    private SKColor? GetCellFillColor(TableCell cell)
    {
        if (cell.Fill is { Type: FillType.Solid, Color: not null })
        {
            return new Color(cell.Fill.Color).AsSkColor();
        }

        return this.GetStyleFill(cell);
    }

    private SKColor? GetStyleFill(TableCell cell)
    {
        var table = (Table)this.Table!;
        var style = (TableStyle)table.TableStyle;

        // Medium Style 2 - Accent 1
        if (style.Guid == MediumStyle2Accent1Guid)
        {
            // Header Row
            if (table.StyleOptions.HasHeaderRow && cell.RowIndex == 0)
            {
                var hex = this.ResolveSchemeColor("accent1");
                return hex != null ? new Color(hex).AsSkColor() : null;
            }
            
            // Banded Rows
            if (table.StyleOptions.HasBandedRows && cell.RowIndex % 2 != 0)
            {
                 // 20% tint of Accent 1
                 var hex = this.ResolveSchemeColor("accent1");
                 if (hex != null)
                 {
                     var color = new Color(hex).AsSkColor();

                     // Approx 20% opacity for simple visual match with the user's expectation
                     return new SKColor(color.Red, color.Green, color.Blue, 51); 
                 }
            }
        }
        
        return null;
    }

    private string? GetStyleFontColorHex(TableCell cell)
    {
        var table = (Table)this.Table!;
        var style = (TableStyle)table.TableStyle;
        
        // Medium Style 2 - Accent 1
        if (style.Guid == MediumStyle2Accent1Guid)
        {
            // Header Row uses white text
            return table.StyleOptions.HasHeaderRow && cell.RowIndex == 0 ? "FFFFFF" : null;
        }
        
        return null;
    }

    private void RenderBorders(SKCanvas canvas, decimal x, decimal y, decimal w, decimal h)
    {
        var table = (Table)this.Table!;
        var style = (TableStyle)table.TableStyle;
        
        // Style borders (Medium Style 2 - Accent 1 implies white borders)
        if (style.Guid != "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
        {
            return;
        }

        using var paint = new SKPaint();
        paint.Color = SKColors.White;
        paint.Style = SKPaintStyle.Stroke;
        paint.StrokeWidth = 1;
        paint.IsAntialias = true;

        var rect = CreateRectFromPoints(x, y, w, h);
        canvas.DrawRect(rect, paint);
    }
}
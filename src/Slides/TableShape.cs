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
        var xPoints = this.X;
        var yPoints = this.Y;
        
        var table = (Table)this.Table!;
        var rowTopPoints = yPoints;
        var rows = table.Rows;
        var columns = table.Columns;

        for (var rowIdx = 0; rowIdx < rows.Count; rowIdx++)
        {
            var row = rows[rowIdx];
            var rowHeightPoints = row.Height;
            var colLeftPoints = xPoints;

            for (var colIdx = 0; colIdx < columns.Count; colIdx++)
            {
                var column = columns[colIdx];
                var colWidthPoints = column.Width;
                var cell = (TableCell)row.Cells[colIdx];

                // Render only if it's the top-left cell of a merge (or single cell)
                if (cell.RowIndex == rowIdx && cell.ColumnIndex == colIdx)
                {
                    int gridSpan = cell.ATableCell.GridSpan?.Value ?? 1;
                    int rowSpan = cell.ATableCell.RowSpan?.Value ?? 1;

                    decimal cellTotalWidth = 0;
                    for (int k = 0; k < gridSpan; k++)
                    {
                        cellTotalWidth += columns[colIdx + k].Width;
                    }

                    decimal cellTotalHeight = 0;
                    for (int k = 0; k < rowSpan; k++)
                    {
                        cellTotalHeight += rows[rowIdx + k].Height;
                    }
                    
                    this.RenderCell(canvas, cell, colLeftPoints, rowTopPoints, cellTotalWidth, cellTotalHeight);
                }

                colLeftPoints += colWidthPoints;
            }

            rowTopPoints += rowHeightPoints;
        }
    }

    private static void RenderCellText(SKCanvas canvas, TableCell cell, decimal x, decimal y, decimal w, decimal h, string? styleFontColorHex)
    {
        var aTextBody = cell.ATableCell.TextBody!;
        
        // Temporarily apply style font color if needed
        var modifiedRunProperties = new System.Collections.Generic.List<(A.RunProperties RunProp, A.SolidFill? OriginalFill)>();
        
        if (styleFontColorHex != null)
        {
            foreach (var aParagraph in aTextBody.Elements<A.Paragraph>())
            {
                foreach (var aRun in aParagraph.Elements<A.Run>())
                {
                    var aRunPr = aRun.GetFirstChild<A.RunProperties>();
                    if (aRunPr == null)
                    {
                        aRunPr = new A.RunProperties();
                        aRun.InsertAt(aRunPr, 0);
                    }
                    
                    var originalFill = aRunPr.GetFirstChild<A.SolidFill>();
                    modifiedRunProperties.Add((aRunPr, originalFill));

                    // Remove existing fill and add style color
                    originalFill?.Remove();
                    var newFill = new A.SolidFill();
                    newFill.Append(new A.RgbColorModelHex { Val = styleFontColorHex });
                    aRunPr.InsertAt(newFill, 0);
                }
            }
        }
        
        try
        {
            var textBoxMargins = new TextBoxMargins(aTextBody);
            var drawingTextBox = new DrawingTextBox(textBoxMargins, aTextBody);
            drawingTextBox.Render(canvas, x, y, w, h);
        }
        finally
        {
            // Restore original state
            if (styleFontColorHex != null)
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
        }
    }
    
    private void RenderCell(SKCanvas canvas, TableCell cell, decimal x, decimal y, decimal w, decimal h)
    {
        // 1. Resolve Fill
        var fillColor = this.GetCellFillColor(cell);
        
        // 2. Render Fill
        if (fillColor != null)
        {
            var rect = new SKRect(
                (float)new Points(x).AsPixels(),
                (float)new Points(y).AsPixels(),
                (float)new Points(x + w).AsPixels(),
                (float)new Points(y + h).AsPixels());

            using var paint = new SKPaint
            {
                Color = fillColor.Value,
                Style = SKPaintStyle.Fill,
                IsAntialias = true
            };
            canvas.DrawRect(rect, paint);
        }

        // 3. Render Borders
        this.RenderBorders(canvas, x, y, w, h);

        // 4. Render Text with style font color
        if (cell.ATableCell.TextBody != null)
        {
            var styleFontColorHex = this.GetStyleFontColorHex(cell);
            RenderCellText(canvas, cell, x, y, w, h, styleFontColorHex);
        }
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
        if (style.Guid == "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
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
        if (style.Guid == "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
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
        
        // Explicit borders (ignoring for now as WIP doesn't use them, but should be here)
        // ...

        // Style borders (Medium Style 2 - Accent 1 implies white borders)
        if (style.Guid == "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
        {
             using var paint = new SKPaint
             {
                 Color = SKColors.White,
                 Style = SKPaintStyle.Stroke,
                 StrokeWidth = 1, 
                 IsAntialias = true
             };
             
             var rect = new SKRect((float)new Points(x).AsPixels(), (float)new Points(y).AsPixels(), (float)new Points(x + w).AsPixels(), (float)new Points(y + h).AsPixels());
             canvas.DrawRect(rect, paint);
        }
    }
}
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Positions;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts;

internal sealed class ChartShape(Chart chartModel, P.GraphicFrame pGraphicFrame) : DrawingShape(new Position(pGraphicFrame),
    new ShapeSize(pGraphicFrame), new ShapeId(pGraphicFrame), pGraphicFrame)
{
    private readonly Chart chart = chartModel;

    public override IBarChart? BarChart => this.IsBarChart(BarDirectionValues.Bar) ? this.chart : null;

    public override IColumnChart? ColumnChart => this.IsBarChart(BarDirectionValues.Column) ? this.chart : null;

    public override ILineChart? LineChart =>
        this.IsChartType(ChartType.LineChart, ChartType.Line3DChart)
        || (this.chart.Type == ChartType.Combination
            && (this.HasChartElement<global::DocumentFormat.OpenXml.Drawing.Charts.LineChart>()
                || this.HasChartElement<global::DocumentFormat.OpenXml.Drawing.Charts.Line3DChart>()))
            ? this.chart
            : null;

    public override IPieChart? PieChart =>
        this.IsChartType(ChartType.PieChart, ChartType.Pie3DChart, ChartType.DoughnutChart)
        || (this.chart.Type == ChartType.Combination
            && (this.HasChartElement<global::DocumentFormat.OpenXml.Drawing.Charts.PieChart>()
                || this.HasChartElement<global::DocumentFormat.OpenXml.Drawing.Charts.Pie3DChart>()
                || this.HasChartElement<global::DocumentFormat.OpenXml.Drawing.Charts.DoughnutChart>()))
            ? this.chart
            : null;

    public override IScatterChart? ScatterChart =>
        this.IsChartType(ChartType.ScatterChart)
        || (this.chart.Type == ChartType.Combination && this.HasChartElement<global::DocumentFormat.OpenXml.Drawing.Charts.ScatterChart>())
            ? this.chart
            : null;

    public override IBubbleChart? BubbleChart =>
        this.IsChartType(ChartType.BubbleChart)
        || (this.chart.Type == ChartType.Combination && this.HasChartElement<global::DocumentFormat.OpenXml.Drawing.Charts.BubbleChart>())
            ? this.chart
            : null;

    public override IAreaChart? AreaChart =>
        this.IsChartType(ChartType.AreaChart, ChartType.Area3DChart)
        || (this.chart.Type == ChartType.Combination
            && (this.HasChartElement<global::DocumentFormat.OpenXml.Drawing.Charts.AreaChart>()
                || this.HasChartElement<global::DocumentFormat.OpenXml.Drawing.Charts.Area3DChart>()))
            ? this.chart
            : null;

    public override ShapeContentType ContentType => ShapeContentType.Chart;

    public override Geometry GeometryType
    {
        get => Geometry.Rectangle;
        set => throw new SCException("Geometry type cannot be set for Chart shape.");
    }

    /// <inheritdoc/>
    public override void CopyTo(P.ShapeTree pShapeTree)
    {
        var pGraphicFrame = (P.GraphicFrame)this.PShapeTreeElement;

        // Clone the graphic frame and add it to the target shape tree
        new SCPShapeTree(pShapeTree).Add(pGraphicFrame);

        // Get the source and target parts
        var sourceOpenXmlPart = pGraphicFrame.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var targetOpenXmlPart = pShapeTree.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;

        // If source and target parts are the same, no need to copy chart parts
        if (sourceOpenXmlPart == targetOpenXmlPart)
        {
            return;
        }

        // Get the source chart part via the chart reference
        var sourceChartReference = pGraphicFrame.Graphic?.GraphicData?.GetFirstChild<ChartReference>();
        if (sourceChartReference?.Id == null)
        {
            return;
        }

        var sourceChartPart = (ChartPart)sourceOpenXmlPart.GetPartById(sourceChartReference.Id!);

        // Create a new chart part in the target slide
        var targetChartPartRId = new SCOpenXmlPart(targetOpenXmlPart).NextRelationshipId();
        var targetChartPart = targetOpenXmlPart.AddNewPart<ChartPart>(sourceChartPart.ContentType, targetChartPartRId);

        // Copy the chart content by cloning the ChartSpace DOM
        targetChartPart.ChartSpace = (ChartSpace)sourceChartPart.ChartSpace!.CloneNode(true);

        // Copy all child parts (embedded workbook, chart styles, color styles, images, etc.)
        CopyChartChildParts(sourceChartPart, targetChartPart);

        // Update the copied graphic frame with the new chart reference ID
        var copiedGraphicFrame = pShapeTree.Elements<P.GraphicFrame>().Last();
        var copiedChartReference = copiedGraphicFrame.Graphic!.GraphicData!.GetFirstChild<ChartReference>()!;
        copiedChartReference.Id = targetChartPartRId;
    }

    internal override void Render(SKCanvas canvas)
    {
        if (this.chart.Type == ChartType.PieChart)
        {
            this.RenderPieChart(canvas);
        }
        else if (this.chart.Type == ChartType.BarChart)
        {
            this.RenderBarChart(canvas);
        }
        else
        {
            // For other chart types, render as a placeholder rectangle
            // Charts use GraphicFrame which lacks ShapeProperties, so we render directly
            this.RenderChartPlaceholder(canvas);
        }
    }

    private static SKColor[] GetPieChartColors()
    {
        return
        [
            SKColor.Parse("#4472C4"), // Blue
            SKColor.Parse("#ED7D31"), // Orange
            SKColor.Parse("#A5A5A5"), // Gray
            SKColor.Parse("#FFC000"), // Yellow
            SKColor.Parse("#5B9BD5"), // Light Blue
            SKColor.Parse("#70AD47") // Green
        ];
    }

    private static ChartLayout CalculateChartLayout(ChartBounds bounds, PieChartData chartData)
    {
        var titleHeight = string.IsNullOrEmpty(chartData.Title) ? 0f : 40f;
        var legendWidth = chartData.ShowLegend ? 150f : 0f;
        var availableWidth = bounds.Width - legendWidth;
        var availableHeight = bounds.Height - titleHeight;
        var pieSize = Math.Min(availableWidth, availableHeight) * 0.75f;
        var centerX = bounds.X + (availableWidth / 2);
        var centerY = bounds.Y + titleHeight + (availableHeight / 2);
        var radius = pieSize / 2;

        return new ChartLayout(centerX, centerY, radius, availableWidth);
    }

    private static void DrawTitle(SKCanvas canvas, string title, ChartBounds bounds)
    {
        if (string.IsNullOrEmpty(title))
        {
            return;
        }

        using var titleFont = new SKFont(SKTypeface.FromFamilyName("Arial", SKFontStyle.Bold), 18);
        using var titlePaint = new SKPaint();
        titlePaint.Color = SKColors.Black;
        titlePaint.IsAntialias = true;
        var titleWidth = titleFont.MeasureText(title);
        canvas.DrawText(
            title,
            bounds.X + ((bounds.Width - titleWidth) / 2),
            bounds.Y + 25,
            SKTextAlign.Left,
            titleFont,
            titlePaint);
    }

    private static List<SliceAngle> DrawPieSlices(
        SKCanvas canvas,
        PieChartData chartData,
        double total,
        ChartLayout layout,
        SKColor[] colors)
    {
        var startAngle = -90f;
        var sliceAngles = new List<SliceAngle>();

        for (var i = 0; i < chartData.Values.Count; i++)
        {
            var sweepAngle = (float)(chartData.Values[i] / total * 360);
            DrawPieSlice(canvas, layout, startAngle, sweepAngle, colors[i % colors.Length]);
            sliceAngles.Add(new SliceAngle(startAngle, sweepAngle, i));
            startAngle += sweepAngle;
        }

        return sliceAngles;
    }

    private static void DrawPieSlice(
        SKCanvas canvas,
        ChartLayout layout,
        float startAngle,
        float sweepAngle,
        SKColor color)
    {
        using var paint = new SKPaint();
        paint.Color = color;
        paint.Style = SKPaintStyle.Fill;
        paint.IsAntialias = true;

        var rect = new SKRect(
            layout.CenterX - layout.Radius,
            layout.CenterY - layout.Radius,
            layout.CenterX + layout.Radius,
            layout.CenterY + layout.Radius);

        using var path = new SKPath();
        path.MoveTo(layout.CenterX, layout.CenterY);
        path.ArcTo(rect, startAngle, sweepAngle, false);
        path.Close();

        canvas.DrawPath(path, paint);

        using var outlinePaint = new SKPaint();
        outlinePaint.Color = SKColors.White;
        outlinePaint.Style = SKPaintStyle.Stroke;
        outlinePaint.StrokeWidth = 2;
        outlinePaint.IsAntialias = true;
        canvas.DrawPath(path, outlinePaint);
    }

    private static void DrawDataLabels(
        SKCanvas canvas,
        PieChartData chartData,
        List<SliceAngle> sliceAngles,
        ChartLayout layout)
    {
        if (!chartData.ShowDataLabels)
        {
            return;
        }

        using var labelFont = new SKFont(SKTypeface.FromFamilyName("Arial"));
        using var labelPaint = new SKPaint();
        labelPaint.Color = SKColors.Black;
        labelPaint.IsAntialias = true;

        foreach (var slice in sliceAngles)
        {
            var angle = slice.Start + (slice.Sweep / 2);
            var angleRad = angle * Math.PI / 180;
            var labelRadius = layout.Radius * 0.6f;
            var labelX = layout.CenterX + (labelRadius * (float)Math.Cos(angleRad));
            var labelY = layout.CenterY + (labelRadius * (float)Math.Sin(angleRad));

            var labelText = chartData.Values[slice.Index].ToString(CultureInfo.InvariantCulture);
            var textWidth = labelFont.MeasureText(labelText);
            canvas.DrawText(labelText, labelX - (textWidth / 2), labelY + 4, SKTextAlign.Left, labelFont, labelPaint);
        }
    }

    private static void DrawLegend(SKCanvas canvas, PieChartData chartData, ChartLayout layout, SKColor[] colors)
    {
        if (!chartData.ShowLegend || chartData.Categories.Count == 0)
        {
            return;
        }

        const float legendItemHeight = 25f;
        var legendX = layout.CenterX + (layout.AvailableWidth / 2) + 20;
        var totalLegendHeight = chartData.Categories.Count * legendItemHeight;
        var legendY = layout.CenterY - (totalLegendHeight / 2);

        using var legendFont = new SKFont(SKTypeface.FromFamilyName("Arial"), 11);
        using var legendTextPaint = new SKPaint();
        legendTextPaint.Color = SKColors.Black;
        legendTextPaint.IsAntialias = true;

        for (var i = 0; i < chartData.Categories.Count; i++)
        {
            var itemY = legendY + (i * legendItemHeight);
            DrawLegendItem(
                canvas,
                legendX,
                itemY,
                chartData.Categories[i],
                colors[i % colors.Length],
                legendFont,
                legendTextPaint);
        }
    }

    private static void DrawLegendItem(
        SKCanvas canvas,
        float legendX,
        float itemY,
        string category,
        SKColor color,
        SKFont font,
        SKPaint textPaint)
    {
        using var boxPaint = new SKPaint();
        boxPaint.Color = color;
        boxPaint.Style = SKPaintStyle.Fill;
        boxPaint.IsAntialias = true;
        var boxRect = new SKRect(legendX, itemY, legendX + 12, itemY + 12);
        canvas.DrawRect(boxRect, boxPaint);
        canvas.DrawText(category, legendX + 18, itemY + 10, SKTextAlign.Left, font, textPaint);
    }

    private static void ExtractTitle(
        DocumentFormat.OpenXml.Drawing.Charts.Chart? chart,
        PieChartSeries pieChartSeries,
        PieChartData data)
    {
        var autoTitleDeleted = chart?.AutoTitleDeleted?.Val?.Value ?? false;
        if (autoTitleDeleted)
        {
            return;
        }

        var title = chart?.Title;
        data.Title = title != null
            ? GetTitleFromChartTitle(title)
            : GetTitleFromSeriesName(pieChartSeries);
    }

    private static string GetTitleFromChartTitle(Title title)
    {
        var chartText = title.ChartText;
        var richText = chartText?.RichText;
        if (richText != null)
        {
            return GetTitleFromRichText(richText);
        }

        return GetTitleFromStringCache(chartText);
    }

    private static string GetTitleFromRichText(DocumentFormat.OpenXml.OpenXmlElement richText)
    {
        return string.Concat(
            richText
                .Descendants<DocumentFormat.OpenXml.Drawing.Run>()
                .Select(r => r.Text?.Text ?? string.Empty));
    }

    private static string GetTitleFromStringCache(ChartText? chartText)
    {
        return chartText?.StringReference?.StringCache?.Elements<StringPoint>().FirstOrDefault()?.NumericValue?.Text ??
               string.Empty;
    }

    private static string GetTitleFromSeriesName(PieChartSeries pieChartSeries)
    {
        var seriesText = pieChartSeries.GetFirstChild<SeriesText>();
        if (seriesText == null)
        {
            return string.Empty;
        }

        var strRef = seriesText.StringReference;
        var strCache = strRef?.StringCache;
        if (strCache != null)
        {
            var firstPoint = strCache.Elements<StringPoint>().FirstOrDefault();
            return firstPoint?.NumericValue?.Text ?? string.Empty;
        }

        return string.Empty;
    }

    private static void ExtractLegendVisibility(DocumentFormat.OpenXml.Drawing.Charts.Chart? chart, PieChartData data)
    {
        data.ShowLegend = chart?.Legend != null;
    }

    private static void ExtractDataLabelsVisibility(ChartSpace chartSpace, PieChartData data)
    {
        var pieChart = chartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.PieChart>().FirstOrDefault();
        var dataLabels = pieChart?.Descendants<DataLabels>().FirstOrDefault();
        if (dataLabels != null)
        {
            var showValue = dataLabels.GetFirstChild<ShowValue>();
            data.ShowDataLabels = showValue?.Val?.Value ?? false;
        }
    }

    private static void ExtractCategories(PieChartSeries pieChartSeries, PieChartData data)
    {
        var categoryAxisData = pieChartSeries.GetFirstChild<CategoryAxisData>();
        var points = GetCategoryPoints(categoryAxisData);
        if (points == null)
        {
            return;
        }

        foreach (var stringPoint in points.OrderBy(sp => sp.Index?.Value ?? 0))
        {
            data.Categories.Add(stringPoint.NumericValue?.Text ?? string.Empty);
        }
    }

    private static IEnumerable<StringPoint>? GetCategoryPoints(CategoryAxisData? categoryAxisData)
    {
        return categoryAxisData?.StringLiteral?.Elements<StringPoint>()
               ?? categoryAxisData?.StringReference?.StringCache?.Elements<StringPoint>();
    }

    private static void ExtractValues(PieChartSeries pieChartSeries, PieChartData data)
    {
        var values = pieChartSeries.GetFirstChild<Values>();
        if (values == null)
        {
            return;
        }

        var literalPoints = values.NumberLiteral?.Elements<NumericPoint>();
        if (literalPoints != null)
        {
            AddParsedValues(literalPoints, data);
            return;
        }

        var cachedPoints = values.NumberReference?.NumberingCache?.Elements<NumericPoint>();
        if (cachedPoints == null)
        {
            return;
        }

        AddParsedValues(cachedPoints, data);
    }

    private static void AddParsedValues(IEnumerable<NumericPoint> numericPoints, PieChartData data)
    {
        foreach (var numericPoint in numericPoints.OrderBy(np => np.Index?.Value ?? 0))
        {
            if (!double.TryParse(
                    numericPoint.NumericValue?.Text,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out var val))
            {
                continue;
            }

            data.Values.Add(val);
        }
    }

    private static void DrawBarChartAxes(
        SKCanvas canvas,
        float chartAreaX,
        float chartAreaY,
        float chartAreaWidth,
        float chartAreaHeight,
        double maxValue)
    {
        using var axisPaint = new SKPaint();
        axisPaint.Color = SKColors.Black;
        axisPaint.Style = SKPaintStyle.Stroke;
        axisPaint.StrokeWidth = 1;
        axisPaint.IsAntialias = true;

        // Draw Y axis (vertical line on the left)
        canvas.DrawLine(chartAreaX, chartAreaY, chartAreaX, chartAreaY + chartAreaHeight, axisPaint);

        // Draw X axis (horizontal line at the bottom)
        canvas.DrawLine(
            chartAreaX, 
            chartAreaY + chartAreaHeight, 
            chartAreaX + chartAreaWidth,
            chartAreaY + chartAreaHeight, 
            axisPaint);

        // Draw X axis ticks and labels
        using var tickFont = new SKFont(SKTypeface.FromFamilyName("Arial"), 9);
        using var tickPaint = new SKPaint();
        tickPaint.Color = SKColors.Black;
        tickPaint.IsAntialias = true;

        const int tickCount = 6;
        for (var i = 0; i <= tickCount; i++)
        {
            var tickX = chartAreaX + (i * chartAreaWidth / tickCount);
            var tickValue = (maxValue * i / tickCount).ToString("F0", CultureInfo.InvariantCulture);

            canvas.DrawLine(tickX, chartAreaY + chartAreaHeight, tickX, chartAreaY + chartAreaHeight + 5, axisPaint);
            canvas.DrawText(tickValue, tickX, chartAreaY + chartAreaHeight + 15, SKTextAlign.Center, tickFont, tickPaint);
        }
    }

    private static void DrawBarChartLegend(
        SKCanvas canvas,
        ChartBounds bounds,
        ISeriesCollection seriesCollection,
        SKColor[] colors)
    {
        const float legendItemHeight = 20f;
        var legendX = bounds.X + bounds.Width - 90f;
        var legendY = bounds.Y + (bounds.Height / 2) - (seriesCollection.Count * legendItemHeight / 2);

        using var legendFont = new SKFont(SKTypeface.FromFamilyName("Arial"), 10);
        using var legendTextPaint = new SKPaint();
        legendTextPaint.Color = SKColors.Black;
        legendTextPaint.IsAntialias = true;

        for (var i = 0; i < seriesCollection.Count; i++)
        {
            var series = seriesCollection[i];
            var itemY = legendY + (i * legendItemHeight);

            using var boxPaint = new SKPaint();
            boxPaint.Color = colors[i % colors.Length];
            boxPaint.Style = SKPaintStyle.Fill;
            boxPaint.IsAntialias = true;
            canvas.DrawRect(new SKRect(legendX, itemY, legendX + 10, itemY + 10), boxPaint);

            var seriesName = series.HasName ? series.Name : $"Series {i + 1}";
            canvas.DrawText(seriesName, legendX + 15, itemY + 9, SKTextAlign.Left, legendFont, legendTextPaint);
        }
    }

    private static void CopyChartChildParts(ChartPart sourceChartPart, ChartPart targetChartPart)
    {
        foreach (var child in sourceChartPart.Parts)
        {
            var childRelationshipId = child.RelationshipId;
            var childPart = child.OpenXmlPart;
            if (childPart is EmbeddedPackagePart embeddedPackagePart)
            {
                CopyEmbeddedPackagePart(embeddedPackagePart, targetChartPart, childRelationshipId);
            }
            else
            {
                // For other parts (chart styles, color styles, images), add them directly
                targetChartPart.AddPart(childPart, childRelationshipId);
            }
        }
    }

    private static void CopyEmbeddedPackagePart(
        EmbeddedPackagePart sourceEmbeddedPackagePart,
        ChartPart targetChartPart,
        string relationshipId)
    {
        var destinationPart = targetChartPart.AddNewPart<EmbeddedPackagePart>(
            sourceEmbeddedPackagePart.ContentType,
            relationshipId);
        using var sourceStream = sourceEmbeddedPackagePart.GetStream(FileMode.Open);
        sourceStream.Position = 0;
        using var destinationStream = destinationPart.GetStream(FileMode.Create, FileAccess.Write);
        sourceStream.CopyTo(destinationStream);
    }

    private bool IsChartType(params ChartType[] types) => types.Contains(this.chart.Type);

    private bool HasChartElement<T>()
        where T : OpenXmlElement
    {
        var chartPart = this.GetChartPart();
        var chartSpace = chartPart?.ChartSpace;
        if (chartSpace == null)
        {
            return false;
        }

        return chartSpace.Descendants<T>().Any();
    }

    private bool IsBarChart(BarDirectionValues direction)
    {
        if (!this.IsChartType(ChartType.BarChart, ChartType.Bar3DChart, ChartType.Combination))
        {
            return false;
        }

        var barDirection = this.GetBarDirection();
        if (barDirection == null)
        {
            return this.IsChartType(ChartType.BarChart, ChartType.Bar3DChart) && direction == BarDirectionValues.Column;
        }

        return barDirection == direction;
    }

    private BarDirectionValues? GetBarDirection()
    {
        var chartPart = this.GetChartPart();
        if (chartPart?.ChartSpace == null)
        {
            return null;
        }

        var barChart = chartPart.ChartSpace
            .Descendants<DocumentFormat.OpenXml.Drawing.Charts.BarChart>()
            .FirstOrDefault();
        var barChartDirection = barChart?.BarDirection?.Val?.Value;
        if (barChartDirection != null)
        {
            return barChartDirection;
        }

        var bar3DChart = chartPart.ChartSpace
            .Descendants<DocumentFormat.OpenXml.Drawing.Charts.Bar3DChart>()
            .FirstOrDefault();
        return bar3DChart?.BarDirection?.Val?.Value;
    }

    private void RenderChartPlaceholder(SKCanvas canvas)
    {
        var x = (float)new Units.Points(this.X).AsPixels();
        var y = (float)new Units.Points(this.Y).AsPixels();
        var width = (float)new Units.Points(this.Width).AsPixels();
        var height = (float)new Units.Points(this.Height).AsPixels();
        var rect = new SKRect(x, y, x + width, y + height);

        using var fillPaint = new SKPaint();
        fillPaint.Color = SKColors.White;
        fillPaint.Style = SKPaintStyle.Fill;
        fillPaint.IsAntialias = true;
        canvas.DrawRect(rect, fillPaint);

        using var outlinePaint = new SKPaint();
        outlinePaint.Color = SKColors.LightGray;
        outlinePaint.Style = SKPaintStyle.Stroke;
        outlinePaint.StrokeWidth = 1;
        outlinePaint.IsAntialias = true;
        canvas.DrawRect(rect, outlinePaint);
    }

    private void RenderBarChart(SKCanvas canvas)
    {
        var chart = this.chart;

        var bounds = this.CalculateChartBounds();
        var colors = GetPieChartColors();

        // Draw a white background
        using var bgPaint = new SKPaint { Color = SKColors.White, Style = SKPaintStyle.Fill };
        canvas.DrawRect(new SKRect(bounds.X, bounds.Y, bounds.X + bounds.Width, bounds.Y + bounds.Height), bgPaint);

        var categories = chart.Categories;
        var seriesCollection = chart.SeriesCollection;
        if (categories == null || categories.Count == 0 || seriesCollection.Count == 0)
        {
            return;
        }

        // Calculate chart area (leave margins for axes and legend)
        const float leftMargin = 80f;
        const float rightMargin = 100f;
        const float topMargin = 20f;
        const float bottomMargin = 30f;

        var chartAreaX = bounds.X + leftMargin;
        var chartAreaY = bounds.Y + topMargin;
        var chartAreaWidth = bounds.Width - leftMargin - rightMargin;
        var chartAreaHeight = bounds.Height - topMargin - bottomMargin;

        // Find max value for scaling
        var maxValue = seriesCollection
            .SelectMany(series => series.Points)
            .Select(point => point.Value)
            .Where(value => value > 0)
            .DefaultIfEmpty(0.0)
            .Max();

        DrawBarChartAxes(canvas, chartAreaX, chartAreaY, chartAreaWidth, chartAreaHeight, maxValue);

        // Draw bars (horizontal for bar chart)
        var categoryCount = categories.Count;
        var seriesCount = seriesCollection.Count;
        var categoryHeight = chartAreaHeight / categoryCount;
        var barHeight = (categoryHeight * 0.7f) / seriesCount;
        var categoryPadding = categoryHeight * 0.15f;

        for (var catIndex = 0; catIndex < categoryCount; catIndex++)
        {
            var categoryY = chartAreaY + (catIndex * categoryHeight) + categoryPadding;

            for (var serIndex = 0; serIndex < seriesCount; serIndex++)
            {
                var series = seriesCollection[serIndex];
                if (catIndex >= series.Points.Count)
                {
                    continue;
                }

                var value = series.Points[catIndex].Value;
                var barWidth = (float)(value / maxValue * chartAreaWidth);
                var barY = categoryY + (serIndex * barHeight);

                using var barPaint = new SKPaint
                {
                    Color = colors[serIndex % colors.Length], Style = SKPaintStyle.Fill, IsAntialias = true
                };
                canvas.DrawRect(new SKRect(chartAreaX, barY, chartAreaX + barWidth, barY + barHeight), barPaint);
            }

            // Draw category label
            using var labelFont = new SKFont(SKTypeface.FromFamilyName("Arial"), 10);
            using var labelPaint = new SKPaint { Color = SKColors.Black, IsAntialias = true };
            var categoryName = categories[catIndex].Name;
            var labelY = categoryY + (categoryHeight * 0.35f);
            canvas.DrawText(categoryName, bounds.X + 5, labelY + 4, SKTextAlign.Left, labelFont, labelPaint);
        }

        // Draw legend
        DrawBarChartLegend(canvas, bounds, seriesCollection, colors);
    }

    private void RenderPieChart(SKCanvas canvas)
    {
        var chartData = this.GetPieChartData();
        if (chartData == null || chartData.Values.Count == 0)
        {
            return;
        }

        var total = chartData.Values.Sum();
        if (total <= 0)
        {
            return;
        }

        var bounds = this.CalculateChartBounds();
        var layout = CalculateChartLayout(bounds, chartData);
        var colors = GetPieChartColors();

        DrawTitle(canvas, chartData.Title, bounds);
        var sliceAngles = DrawPieSlices(canvas, chartData, total, layout, colors);
        DrawDataLabels(canvas, chartData, sliceAngles, layout);
        DrawLegend(canvas, chartData, layout, colors);
    }

    private ChartBounds CalculateChartBounds()
    {
        return new ChartBounds(
            (float)new Units.Points(this.X).AsPixels(),
            (float)new Units.Points(this.Y).AsPixels(),
            (float)new Units.Points(this.Width).AsPixels(),
            (float)new Units.Points(this.Height).AsPixels());
    }

    private PieChartData? GetPieChartData()
    {
        var chartPart = this.GetChartPart();
        if (chartPart == null)
        {
            return null;
        }

        var chartSpace = chartPart.ChartSpace;

        var pieChartSeries = chartSpace!.Descendants<PieChartSeries>().FirstOrDefault();
        if (pieChartSeries == null)
        {
            return null;
        }

        var data = new PieChartData();
        var chart = chartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();

        ExtractTitle(chart, pieChartSeries, data);
        ExtractLegendVisibility(chart, data);
        ExtractDataLabelsVisibility(chartSpace, data);
        ExtractCategories(pieChartSeries, data);
        ExtractValues(pieChartSeries, data);

        return data.Values.Count > 0 ? data : null;
    }

    private ChartPart? GetChartPart()
    {
        var pGraphicFrame = (P.GraphicFrame)this.PShapeTreeElement;
        var graphicData = pGraphicFrame.Graphic?.GraphicData;
        if (graphicData == null)
        {
            return null;
        }

        var chartReference = graphicData.GetFirstChild<ChartReference>();
        if (chartReference?.Id == null)
        {
            return null;
        }

        var slidePart = pGraphicFrame.Ancestors<P.Slide>().FirstOrDefault()?.SlidePart;
        if (slidePart == null)
        {
            return null;
        }

        return (ChartPart)slidePart.GetPartById(chartReference.Id!);
    }
}
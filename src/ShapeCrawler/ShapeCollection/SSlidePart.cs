namespace ShapeCrawler.ShapeCollection;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

internal readonly ref struct SSlidePart(SlidePart slidePart)
{
    internal void AddPieChart()
    {
        var chartPart = slidePart.AddNewPart<ChartPart>("rId3");
        GeneratePieChartContent(chartPart);
        this.InsertChartGraphicFrame(chartPart);
    }
    
    /// <summary>
    /// Generates the XML for a simple Pie Chart with 3 categories.
    /// </summary>
    private static void GeneratePieChartContent(ChartPart chartPart)
    {
        // Create the ChartSpace
        ChartSpace chartSpace = new ChartSpace();
        chartSpace.Append(new EditingLanguage() { Val = "en-US" });
        chartSpace.Append(new RoundedCorners() { Val = false });

        // Create the Chart
        Chart chart = new Chart();
        chart.Append(new AutoTitleDeleted() { Val = true }); // Hide default chart title

        PlotArea plotArea = new PlotArea();
        plotArea.Append(new Layout());

        // Create the PieChart element
        PieChart pieChart = new PieChart();
        pieChart.Append(new VaryColors() { Val = true });

        // PieChartSeries: define series index, order, and label
        PieChartSeries series = new PieChartSeries(
            new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = 0 },
            new Order() { Val = 0 },
            new SeriesText(new NumericValue() { Text = "Sample Series" }));

        // --- Categories ---
        var stringLiteral = new StringLiteral { PointCount = new PointCount() { Val = 3 } };
        stringLiteral.Append(new StringPoint() { Index = 0, NumericValue = new NumericValue("Category A") });
        stringLiteral.Append(new StringPoint() { Index = 1, NumericValue = new NumericValue("Category B") });
        stringLiteral.Append(new StringPoint() { Index = 2, NumericValue = new NumericValue("Category C") });

        CategoryAxisData categoryAxisData = new CategoryAxisData();
        categoryAxisData.Append(stringLiteral);

        // --- Values ---
        var numberLiteral = new NumberLiteral { FormatCode = new FormatCode("General"), PointCount = new PointCount() { Val = 3 } };
        numberLiteral.Append(new NumericPoint() { Index = 0, NumericValue = new NumericValue("10") });
        numberLiteral.Append(new NumericPoint() { Index = 1, NumericValue = new NumericValue("30") });
        numberLiteral.Append(new NumericPoint() { Index = 2, NumericValue = new NumericValue("60") });

        Values values = new Values();
        values.Append(numberLiteral);

        // Append categories and values to the series
        series.Append(categoryAxisData);
        series.Append(values);

        // Add series to the PieChart
        pieChart.Append(series);

        // Add the PieChart to the PlotArea
        plotArea.Append(pieChart);

        // Complete the chart
        chart.Append(plotArea);

        // Add the chart to the ChartSpace
        chartSpace.Append(chart);

        // Assign ChartSpace to the ChartPart
        chartPart.ChartSpace = chartSpace;
    }

    /// <summary>
    /// Inserts a graphic frame in the slide that references the chart part by ID.
    /// </summary>
    private void InsertChartGraphicFrame(ChartPart chartPart)
    {
        // Retrieve the shape tree of the slide
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree;

        // Create a new GraphicFrame
        var graphicFrame = new GraphicFrame
        {
            // Give it an ID and name
            NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties() { Id = 2U, Name = "PieChart" },
                new NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            
            // Position/Size of the chart in EMUs (English Metric Units)
            // For reference: 914400 EMUs = 1 inch
            Transform = new Transform(
                new A.Offset() { X = 1524000L, Y = 1524000L }, // 1.67in from top-left corner
                new A.Extents() { Cx = 6096000L, Cy = 3429000L }), // 6.67in wide, 3.75in high
            Graphic = new A.Graphic(
                new A.GraphicData(
                    new ChartReference() { Id = slidePart.GetIdOfPart(chartPart) }) // "rId3"
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })
        };

        // Append the graphic frame to the shape tree
        shapeTree!.Append(graphicFrame);
    }
}
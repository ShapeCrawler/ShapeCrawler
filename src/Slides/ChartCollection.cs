using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Charts;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Slides;

internal sealed class ChartCollection(ISlideShapeCollection shapes, SlidePart slidePart)
    : ISlideShapeCollection
{
    public int Count => shapes.Count;

    public IShape this[int index] => shapes[index];
    
    public void AddPieChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName)
    {
        this.AddPieChart(x, y, width, height, categoryValues, seriesName, "Pie Chart");
    }

    public void AddPieChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName,
        string chartName)
    {
        var rId = new SCOpenXmlPart(slidePart).NextRelationshipId();
        var chartPart = slidePart.AddNewPart<ChartPart>(rId);
        new PieChartContent(chartPart, categoryValues, seriesName).Generate();
        this.InsertChartGraphicFrame(chartPart, x, y, width, height, chartName);
    }

    public void AddBarChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName)
    {
        var rId = new SCOpenXmlPart(slidePart).NextRelationshipId();
        var chartPart = slidePart.AddNewPart<ChartPart>(rId);
        new BarChartContent(chartPart, categoryValues, seriesName).Generate();
        this.InsertChartGraphicFrame(chartPart, x, y, width, height, "Bar Chart");
    }

    public void AddScatterChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<double, double> pointValues,
        string seriesName)
    {
        var rId = new SCOpenXmlPart(slidePart).NextRelationshipId();
        var chartPart = slidePart.AddNewPart<ChartPart>(rId);
        new ScatterChartContent(chartPart, pointValues, seriesName).Generate();
        this.InsertChartGraphicFrame(chartPart, x, y, width, height, "Scatter Chart");
    }

    public void AddStackedColumnChart(
        int x,
        int y,
        int width,
        int height,
        IDictionary<string, IList<double>> categoryValues,
        IList<string> seriesNames)
    {
        var rId = new SCOpenXmlPart(slidePart).NextRelationshipId();
        var chartPart = slidePart.AddNewPart<ChartPart>(rId);
        new StackedColumnChartContent(chartPart, categoryValues, seriesNames).Generate();
        this.InsertChartGraphicFrame(chartPart, x, y, width, height, "Stacked Column Chart");
    }

    #region Shapes

    public void Add(IShape addingShape) => shapes.Add(addingShape);

    public void AddAudio(int x, int y, Stream audio) => shapes.AddAudio(x, y, audio);

    public void AddAudio(int x, int y, Stream audio, AudioType type) => shapes.AddAudio(x, y, audio, type);

    public void AddVideo(int x, int y, Stream stream) => shapes.AddVideo(x, y, stream);

    public void AddShape(
        int x,
        int y,
        int width,
        int height,
        Geometry geometry = Geometry.Rectangle
    ) => shapes.AddShape(x, y, width, height, geometry);

    public void AddShape(
        int x,
        int y,
        int width,
        int height,
        Geometry geometry,
        string text
    ) => shapes.AddShape(x, y, width, height, geometry, text);

    public void AddLine(string xml) => shapes.AddLine(xml);

    public void AddLine(
        int startPointX,
        int startPointY,
        int endPointX,
        int endPointY
    ) => shapes.AddLine(startPointX, startPointY, endPointX, endPointY);

    public void AddTable(
        int x,
        int y,
        int columnsCount,
        int rowsCount
    ) => shapes.AddTable(x, y, columnsCount, rowsCount);

    public void AddTable(
        int x,
        int y,
        int columnsCount,
        int rowsCount,
        ITableStyle style
    ) => shapes.AddTable(x, y, columnsCount, rowsCount, style);

    public void AddPicture(Stream imageStream) => shapes.AddPicture(imageStream);

    public IShape AddSmartArt(
        int x,
        int y,
        int width,
        int height,
        SmartArtType smartArtType
    ) => shapes.AddSmartArt(x, y, width, height, smartArtType);

    public IShape Group(IShape[] groupingShapes)
    {
        throw new System.NotImplementedException();
    }

    public IEnumerator<IShape> GetEnumerator() => shapes.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => shapes.GetEnumerator();

    public IShape GetById(int id) => shapes.GetById(id);

    public T GetById<T>(int id)
        where T : IShape => shapes.GetById<T>(id);

    public IShape Shape(string name) => shapes.Shape(name);

    public T Shape<T>(string name)
        where T : IShape => shapes.Shape<T>(name);

    public T Last<T>()
        where T : IShape => shapes.Last<T>();

    #endregion Shapes

    private void InsertChartGraphicFrame(ChartPart chartPart, int x, int y, int width, int height, string chartName)
    {
        var graphicFrame = new GraphicFrame
        {
            NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties { Id = this.GetNextShapeId(), Name = chartName },
                new NonVisualGraphicFrameDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()
            ),
            Transform = new Transform(
                new A.Offset { X = new Points(x).AsEmus(), Y = new Points(y).AsEmus() },
                new A.Extents { Cx = new Points(width).AsEmus(), Cy = new Points(height).AsEmus() }
            ),
            Graphic = new A.Graphic(
                new A.GraphicData(
                    new ChartReference { Id = slidePart.GetIdOfPart(chartPart) })
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                }
            )
        };

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(graphicFrame);
    }

    private uint GetNextShapeId()
    {
        var shapeIds = slidePart.Slide.Descendants<NonVisualDrawingProperties>()
            .Select(p => p.Id?.Value ?? 0)
            .ToList();

        return shapeIds.Count > 0 ? shapeIds.Max() + 1 : 1;
    }
}
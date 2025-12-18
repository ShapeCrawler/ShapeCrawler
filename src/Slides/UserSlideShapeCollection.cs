// ReSharper disable InconsistentNaming
// ReSharper disable UseObjectOrCollectionInitializer

using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Assets;
using ShapeCrawler.Charts;
using ShapeCrawler.Drawing;
using ShapeCrawler.Groups;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tables;
using ShapeCrawler.Units;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class UserSlideShapeCollection : IUserSlideShapeCollection
{
    private readonly ShapeCollection shapes;
    private readonly PictureShapeCollection pictureShapes;
    private readonly AudioVideoShapeCollection audioVideoShapes;
    private readonly ChartShapeCollection chartShapes;
    private readonly FooterPlaceholderShapeCollection footerPlaceholderShapeCollection;
    
    private readonly NewShapeProperties newShapeProperties;
    private readonly SlidePart slidePart;

    internal UserSlideShapeCollection(
        ShapeCollection shapes,
        PictureShapeCollection pictureShapes,
        AudioVideoShapeCollection audioVideoShapes,
        ChartShapeCollection chartShapes,
        SlidePart slidePart)
    {
        this.shapes = shapes;
        this.pictureShapes = pictureShapes;
        this.audioVideoShapes = audioVideoShapes;
        this.chartShapes = chartShapes;
        this.slidePart = slidePart;
        this.newShapeProperties = new NewShapeProperties(this);
        this.footerPlaceholderShapeCollection = new FooterPlaceholderShapeCollection(this, slidePart);
    }

    public int Count => this.shapes.Count;

    public IShape this[int index] => this.shapes[index];

    public void Add(IShape addingShape)
    {
        var pShapeTree = this.slidePart.Slide.CommonSlideData!.ShapeTree!;
        switch (addingShape)
        {
            case PictureShape picture:
                picture.CopyTo(pShapeTree);
                break;
            case TextShape textShape:
                textShape.CopyTo(pShapeTree);
                break;
            case TableShape table:
                table.CopyTo(pShapeTree);
                break;
            case Shape shape:
                shape.CopyTo(pShapeTree);
                break;
            default:
                throw new SCException("Unsupported shape type for adding.");
        }
    }

    public void AddAudio(int x, int y, Stream audio) => this.audioVideoShapes.AddAudio(x, y, audio);

    public void AddAudio(int x, int y, Stream audio, AudioType type) =>
        this.audioVideoShapes.AddAudio(x, y, audio, type);

    public void AddVideo(int x, int y, Stream stream) => this.audioVideoShapes.AddVideo(x, y, stream);

    public void AddPicture(Stream imageStream) => this.pictureShapes.AddPicture(imageStream);

    public void AddPieChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName
    ) => this.chartShapes.AddPieChart(x, y, width, height, categoryValues, seriesName);

    public void AddPieChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName,
        string chartName
    ) => this.chartShapes.AddPieChart(x, y, width, height, categoryValues, seriesName, chartName);

    public void AddBarChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName
    ) => this.chartShapes.AddBarChart(x, y, width, height, categoryValues, seriesName);

    public void AddScatterChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<double, double> pointValues,
        string seriesName
    ) => this.chartShapes.AddScatterChart(x, y, width, height, pointValues, seriesName);

    public void AddStackedColumnChart(
        int x,
        int y,
        int width,
        int height,
        IDictionary<string, IList<double>> categoryValues,
        IList<string> seriesNames
    ) => this.chartShapes.AddStackedColumnChart(x, y, width, height, categoryValues, seriesNames);

    public void AddClusteredBarChart(
        int x,
        int y,
        int width,
        int height,
        IList<string> categories,
        IList<Presentations.DraftChart.SeriesData> seriesData,
        string chartName
    ) => this.chartShapes.AddClusteredBarChart(x, y, width, height, categories, seriesData, chartName);

    public IShape AddSmartArt(
        int x,
        int y,
        int width,
        int height,
        SmartArtType smartArtType)
        => new SCSlidePart(this.slidePart).AddSmartArt(x, y, width, height, smartArtType);

    public IShape Group(IShape[] groupingShapes) =>
        new GroupShape(new P.GroupShape(), groupingShapes, this.newShapeProperties, this.slidePart);

    public void AddShape(int x, int y, int width, int height) => AddShape(x, y, width, height, Geometry.Rectangle);

    public void AddShape(int x, int y, int width, int height, Geometry geometry)
    {
        var xml = new AssetCollection(Assembly.GetExecutingAssembly()).StringOf("new rectangle.xml");
        var pShape = new P.Shape(xml);
        var nextShapeId = this.newShapeProperties.Id();
        this.slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pShape);

        var addedShape = this.shapes.Last<TextShape>();
        addedShape.Name = geometry.ToString();
        addedShape.X = x;
        addedShape.Y = y;
        addedShape.Width = width;
        addedShape.Height = height;
        addedShape.Id = nextShapeId;
        addedShape.GeometryType = geometry;
    }

    public void AddShape(int x, int y, int width, int height, Geometry geometry, string text)
    {
        // First add the basic shape
        var xml = new AssetCollection(Assembly.GetExecutingAssembly()).StringOf("new rectangle.xml");
        var pShape = new P.Shape(xml);
        var nextShapeId = this.newShapeProperties.Id();
        this.slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pShape);

        var addedShape = this.shapes.Last<TextShape>();
        addedShape.Name = geometry.ToString();
        addedShape.X = x;
        addedShape.Y = y;
        addedShape.Width = width;
        addedShape.Height = height;
        addedShape.Id = nextShapeId;
        addedShape.GeometryType = geometry;
        addedShape.ShapeText.SetText(text);
    }

    public void AddTextBox(int x, int y, int width, int height, string text)
    {
        var xml = new AssetCollection(Assembly.GetExecutingAssembly()).StringOf("new rectangle.xml");
        var pShape = new P.Shape(xml);
        var nextShapeId = this.newShapeProperties.Id();
        this.slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pShape);

        var addedShape = this.shapes.Last<TextShape>();
        addedShape.Name = "Text Box";
        addedShape.X = x;
        addedShape.Y = y;
        addedShape.Width = width;
        addedShape.Height = height;
        addedShape.Id = nextShapeId;

        var pNvSpPr = pShape.NonVisualShapeProperties!.GetFirstChild<P.NonVisualShapeDrawingProperties>()!;
        pNvSpPr.TextBox = true;
        pShape.ShapeStyle = null;

        addedShape.ShapeText.SetText(text);
        foreach (var paragraph in addedShape.ShapeText.Paragraphs)
        {
            paragraph.HorizontalAlignment = TextHorizontalAlignment.Left;
        }
    }

    public void AddLine(string xml)
    {
        var newPConnectionShape = new P.ConnectionShape(xml);

        this.slidePart.Slide.CommonSlideData!.ShapeTree!.Append(newPConnectionShape);
    }

    public void AddLine(int startPointX, int startPointY, int endPointX, int endPointY)
        => new ConnectionShape(this.slidePart, this.newShapeProperties)
            .Create(startPointX, startPointY, endPointX, endPointY);

    public void AddTable(int x, int y, int columnsCount, int rowsCount)
        => this.AddTable(x, y, columnsCount, rowsCount, CommonTableStyles.MediumStyle2Accent1);

    public void AddTable(int x, int y, int columnsCount, int rowsCount, ITableStyle style)
    {
        var shapeName = this.newShapeProperties.TableName();
        var xEmu = new Points(x).AsEmus();
        var yEmu = new Points(y).AsEmus();
        var tableHeightEmu = Constants.DefaultRowHeightEmu * rowsCount;

        var graphicFrame = new P.GraphicFrame();
        var nonVisualGraphicFrameProperties = new P.NonVisualGraphicFrameProperties();
        var nonVisualDrawingProperties = new P.NonVisualDrawingProperties
        {
            Id = (uint)this.newShapeProperties.Id(), Name = shapeName
        };
        var nonVisualGraphicFrameDrawingProperties = new P.NonVisualGraphicFrameDrawingProperties();
        var applicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();
        nonVisualGraphicFrameProperties.Append(nonVisualDrawingProperties);
        nonVisualGraphicFrameProperties.Append(nonVisualGraphicFrameDrawingProperties);
        nonVisualGraphicFrameProperties.Append(applicationNonVisualDrawingProperties);

        const long DefaultTableWidthEMUs = 8_128_000L;
        var offset = new A.Offset { X = xEmu, Y = yEmu };
        var extents = new A.Extents { Cx = DefaultTableWidthEMUs, Cy = tableHeightEmu };
        var pTransform = new P.Transform(offset, extents);

        var graphic = new A.Graphic();
#pragma warning disable S1075 // PowerPoint/Open XML requires this exact schema URI for table graphic data.
        var graphicData = new A.GraphicData { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };
#pragma warning restore S1075
        var aTable = new A.Table();

        var tableProperties = new A.TableProperties { FirstRow = true, BandRow = true };
        var tableStyleId = new A.TableStyleId { Text = ((TableStyle)style).Guid };
        tableProperties.Append(tableStyleId);

        var tableGrid = new A.TableGrid();
        var gridWidthEmu = DefaultTableWidthEMUs / columnsCount;
        for (var i = 0; i < columnsCount; i++)
        {
            var gridColumn = new A.GridColumn { Width = gridWidthEmu };
            tableGrid.Append(gridColumn);
        }

        aTable.Append(tableProperties);
        aTable.Append(tableGrid);
        for (var i = 0; i < rowsCount; i++)
        {
            var aTableRow = new A.TableRow { Height = Constants.DefaultRowHeightEmu };
            for (var i2 = 0; i2 < columnsCount; i2++)
            {
                new SCATableRow(aTableRow).AddNewCell();
            }

            aTable.Append(aTableRow);
        }

        graphicData.Append(aTable);
        graphic.Append(graphicData);
        graphicFrame.Append(nonVisualGraphicFrameProperties);
        graphicFrame.Append(pTransform);
        graphicFrame.Append(graphic);

        this.slidePart.Slide.CommonSlideData!.ShapeTree!.Append(graphicFrame);
    }

    public IShape GetById(int id) => this.shapes.GetById(id);

    public T GetById<T>(int id)
        where T : IShape => this.shapes.GetById<T>(id);

    public T GetByName<T>(string name)
        where T : IShape => this.shapes.Shape<T>(name);

    public T Shape<T>(string name)
        where T : IShape => this.shapes.Shape<T>(name);

    public IShape Shape(string name) => this.shapes.Shape(name);

    public T Last<T>()
        where T : IShape => this.shapes.Last<T>();

    public IEnumerator<IShape> GetEnumerator() => this.shapes.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public IShape AddDateAndTime() => this.footerPlaceholderShapeCollection.AddDateAndTime();

    public IShape AddFooter() => this.footerPlaceholderShapeCollection.AddFooter();

    public IShape AddSlideNumber() => this.footerPlaceholderShapeCollection.AddSlideNumber();

    internal void Render(SKCanvas canvas)
    {
        foreach (var shape in this.shapes.GetInternalShapes())
        {
            if (shape.Hidden)
            {
                continue;
            }

            shape.Render(canvas);
        }
    }
}
// ReSharper disable InconsistentNaming
// ReSharper disable UseObjectOrCollectionInitializer

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Assets;
using ShapeCrawler.Drawing;
using ShapeCrawler.Groups;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tables;
using ShapeCrawler.Units;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class SlideShapeCollection(ISlideShapeCollection shapes, SlidePart slidePart) : ISlideShapeCollection
{
    private const double Epsilon = 1e-6;
    private static readonly Dictionary<string, Func<A.ColorScheme, A.Color2Type?>> SchemeColorSelectors =
        new(StringComparer.Ordinal)
        {
            { "dk1", scheme => scheme.Dark1Color },
            { "lt1", scheme => scheme.Light1Color },
            { "dk2", scheme => scheme.Dark2Color },
            { "lt2", scheme => scheme.Light2Color },
            { "accent1", scheme => scheme.Accent1Color },
            { "accent2", scheme => scheme.Accent2Color },
            { "accent3", scheme => scheme.Accent3Color },
            { "accent4", scheme => scheme.Accent4Color },
            { "accent5", scheme => scheme.Accent5Color },
            { "accent6", scheme => scheme.Accent6Color },
            { "hlink", scheme => scheme.Hyperlink },
            { "folHlink", scheme => scheme.FollowedHyperlinkColor }
        };

    private readonly NewShapeProperties newShapeProperties = new(shapes);
    private readonly PlaceholderShapes placeholderShape = new(shapes, slidePart);
    private readonly ConnectionShape connectionShape = new(slidePart, new NewShapeProperties(shapes));
    private readonly TextDrawing textDrawing = new();
    

    public int Count => shapes.Count;

    public IShape this[int index] => shapes[index];

    public void Add(IShape addingShape)
    {
        var pShapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;
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

    public void AddAudio(int x, int y, Stream audio) => shapes.AddAudio(x, y, audio);

    public void AddAudio(int x, int y, Stream audio, AudioType type) => shapes.AddAudio(x, y, audio, type);

    public void AddVideo(int x, int y, Stream stream) => shapes.AddVideo(x, y, stream);

    public void AddPicture(Stream imageStream) => shapes.AddPicture(imageStream);

    public void AddPieChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName
    ) => shapes.AddPieChart(x, y, width, height, categoryValues, seriesName);

    public void AddPieChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName,
        string chartName
    ) => shapes.AddPieChart(x, y, width, height, categoryValues, seriesName, chartName);

    public void AddBarChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName
    ) => shapes.AddBarChart(x, y, width, height, categoryValues, seriesName);

    public void AddScatterChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<double, double> pointValues,
        string seriesName
    ) => shapes.AddScatterChart(x, y, width, height, pointValues, seriesName);

    public void AddStackedColumnChart(
        int x,
        int y,
        int width,
        int height,
        IDictionary<string, IList<double>> categoryValues,
        IList<string> seriesNames
    ) => shapes.AddStackedColumnChart(x, y, width, height, categoryValues, seriesNames);

    public void AddClusteredBarChart(
        int x,
        int y,
        int width,
        int height,
        IList<string> categories,
        IList<Presentations.DraftChart.SeriesData> seriesData,
        string chartName
    ) => shapes.AddClusteredBarChart(x, y, width, height, categories, seriesData, chartName);

    public IShape AddSmartArt(
        int x,
        int y,
        int width,
        int height,
        SmartArtType smartArtType)
        => new SCSlidePart(slidePart).AddSmartArt(x, y, width, height, smartArtType);

    public IShape Group(IShape[] groupingShapes) =>
        new GroupShape(new P.GroupShape(), groupingShapes, this.newShapeProperties, slidePart);

    public void AddShape(int x, int y, int width, int height, Geometry geometry = Geometry.Rectangle)
    {
        var xml = new AssetCollection(Assembly.GetExecutingAssembly()).StringOf("new rectangle.xml");
        var pShape = new P.Shape(xml);
        var nextShapeId = this.newShapeProperties.Id();
        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pShape);

        var addedShape = shapes.Last<TextShape>();
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
        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pShape);

        var addedShape = shapes.Last<TextShape>();
        addedShape.Name = geometry.ToString();
        addedShape.X = x;
        addedShape.Y = y;
        addedShape.Width = width;
        addedShape.Height = height;
        addedShape.Id = nextShapeId;
        addedShape.GeometryType = geometry;
        addedShape.TextBox.SetText(text);
    }

    public void AddLine(string xml)
    {
        var newPConnectionShape = new P.ConnectionShape(xml);

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(newPConnectionShape);
    }

    public void AddLine(int startPointX, int startPointY, int endPointX, int endPointY)
        => this.connectionShape.Create(startPointX, startPointY, endPointX, endPointY);

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

        const long DefaultTableWidthEMUs = 8128000L;
        var offset = new A.Offset { X = xEmu, Y = yEmu };
        var extents = new A.Extents { Cx = DefaultTableWidthEMUs, Cy = tableHeightEmu };
        var pTransform = new P.Transform(offset, extents);

        var graphic = new A.Graphic();
        var graphicData = new A.GraphicData { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };
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

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(graphicFrame);
    }

    public IShape GetById(int id) => shapes.GetById(id);

    public T GetById<T>(int id)
        where T : IShape => shapes.GetById<T>(id);

    public T GetByName<T>(string name)
        where T : IShape => shapes.Shape<T>(name);

    public T Shape<T>(string name)
        where T : IShape => shapes.Shape<T>(name);

    public IShape Shape(string name) => shapes.Shape(name);

    public T Last<T>()
        where T : IShape => shapes.Last<T>();

    public IEnumerator<IShape> GetEnumerator() => shapes.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public IShape AddDateAndTime() => this.placeholderShape.AddDateAndTime();

    public IShape AddFooter() => this.placeholderShape.AddFooter();

    public IShape AddSlideNumber() => this.placeholderShape.AddSlideNumber();

    internal void Render(SKCanvas canvas)
    {
        foreach (var shape in shapes)
        {
            if (shape.Hidden)
            {
                continue;
            }

            this.RenderShape(canvas, shape);
        }
    }
    
    private static decimal GetShapeOutlineWidth(IShape shape)
    {
        var shapeOutline = shape.Outline;

        // Check for explicit outline weight first
        if (shapeOutline is not null && shapeOutline.Weight > 0)
        {
            return new Points(shapeOutline.Weight).AsPixels();
        }

        // Check for style-based outline width
        var styleWidth = GetStyleOutlineWidth(shape);
        return styleWidth;
    }

    private static SKColor ApplyShade(SKColor color, int shadeValue)
    {
        var shadeFactor = shadeValue / 100_000f;

        return new SKColor(
            (byte)(color.Red * shadeFactor),
            (byte)(color.Green * shadeFactor),
            (byte)(color.Blue * shadeFactor),
            color.Alpha);
    }

    private static decimal GetStyleOutlineWidth(IShape shape)
    {
        if (shape.SDKOpenXmlElement is not P.Shape pShape)
        {
            return 0;
        }

        var style = pShape.ShapeStyle;
        var lineRef = style?.LineReference;
        if (lineRef?.Index is null || lineRef.Index.Value == 0)
        {
            return 0;
        }

        // Default line width based on index (idx="2" typically means ~1.5pt line)
        // This is a simplification - proper implementation would look up theme line styles
        var defaultWidth = lineRef.Index.Value * 0.75m;

        return new Points(defaultWidth).AsPixels();
    }

    private static void ApplyRotation(
        SKCanvas canvas,
        IShape shape,
        decimal x,
        decimal y,
        decimal width,
        decimal height)
    {
        if (Math.Abs(shape.Rotation) > Epsilon)
        {
            var centerX = x + (width / 2);
            var centerY = y + (height / 2);
            canvas.RotateDegrees(
                (float)shape.Rotation,
                (float)new Points(centerX).AsPixels(),
                (float)new Points(centerY).AsPixels()
            );
        }
    }
    

    private static string? GetHexFromColorElement(A.Color2Type colorElement)
    {
        var rgbColor = colorElement.RgbColorModelHex;
        if (rgbColor?.Val?.Value is { } rgb)
        {
            return rgb;
        }

        var sysColor = colorElement.SystemColor;
        return sysColor?.LastColor?.Value;
    }

    private void RenderShape(SKCanvas canvas, IShape shape)
    {
        var geometryType = shape.GeometryType;

        switch (geometryType)
        {
            case Geometry.Rectangle:
            case Geometry.RoundedRectangle:
                RenderRectangle(canvas, shape);
                break;
            case Geometry.Ellipse:
                this.RenderEllipse(canvas, shape);
                break;
            default:
                this.RenderText(canvas, shape);
                return;
        }

        this.RenderText(canvas, shape);
    }

    private void RenderText(SKCanvas canvas, IShape shape)
    {
        if (shape.TextBox is null)
        {
            return;
        }

        canvas.Save();
        ApplyRotation(canvas, shape, shape.X, shape.Y, shape.Width, shape.Height);
        this.textDrawing.Render(canvas, shape);
        canvas.Restore();
    }

    private void RenderRectangle(SKCanvas canvas, IShape shape)
    {
        var x = new Points(shape.X).AsPixels();
        var y = new Points(shape.Y).AsPixels();
        var width = new Points(shape.Width).AsPixels();
        var height = new Points(shape.Height).AsPixels();
        var rect = new SKRect((float)x, (float)y, (float)(x + width), (float)(y + height));

        var cornerRadius = 0m;
        if (shape.GeometryType == Geometry.RoundedRectangle)
        {
            // CornerSize is percentage (0-100), where 100 = half of shortest side
            var shortestSide = Math.Min(width, height);
            cornerRadius = shape.CornerSize / 100m * (shortestSide / 2m);
        }

        canvas.Save();
        ApplyRotation(canvas, shape, shape.X, shape.Y, shape.Width, shape.Height);

        this.RenderFill(canvas, shape, rect, cornerRadius);
        this.RenderOutline(canvas, shape, rect, cornerRadius);

        canvas.Restore();
    }

    private void RenderEllipse(SKCanvas canvas, IShape shape)
    {
        var x = new Points(shape.X).AsPixels();
        var y = new Points(shape.Y).AsPixels();
        var width = new Points(shape.Width).AsPixels();
        var height = new Points(shape.Height).AsPixels();
        var rect = new SKRect((float)x, (float)y, (float)(x + width), (float)(y + height));

        canvas.Save();
        ApplyRotation(canvas, shape, shape.X, shape.Y, shape.Width, shape.Height);

        this.RenderEllipseFill(canvas, shape, rect);
        this.RenderEllipseOutline(canvas, shape, rect);

        canvas.Restore();
    }

    private void RenderFill(SKCanvas canvas, IShape shape, SKRect rect, decimal cornerRadius)
    {
        var fillColor = this.GetShapeFillColor(shape);
        if (fillColor is null)
        {
            return;
        }

        using var fillPaint = new SKPaint();
        fillPaint.Color = fillColor.Value;
        fillPaint.Style = SKPaintStyle.Fill;
        fillPaint.IsAntialias = true;

        if (cornerRadius > 0)
        {
            canvas.DrawRoundRect(rect, (float)cornerRadius, (float)cornerRadius, fillPaint);
        }
        else
        {
            canvas.DrawRect(rect, fillPaint);
        }
    }

    private void RenderOutline(SKCanvas canvas, IShape shape, SKRect rect, decimal cornerRadius)
    {
        var outlineColor = this.GetShapeOutlineColor(shape);
        var strokeWidth = GetShapeOutlineWidth(shape);

        if (outlineColor is null || strokeWidth <= 0)
        {
            return;
        }

        using var outlinePaint = new SKPaint();
        outlinePaint.Color = outlineColor.Value;
        outlinePaint.Style = SKPaintStyle.Stroke;
        outlinePaint.StrokeWidth = (float)strokeWidth;
        outlinePaint.IsAntialias = true;

        if (cornerRadius > 0)
        {
            canvas.DrawRoundRect(rect, (float)cornerRadius, (float)cornerRadius, outlinePaint);
        }
        else
        {
            canvas.DrawRect(rect, outlinePaint);
        }
    }

    private void RenderEllipseFill(SKCanvas canvas, IShape shape, SKRect rect)
    {
        var fillColor = this.GetShapeFillColor(shape);
        if (fillColor is null)
        {
            return;
        }

        using var fillPaint = new SKPaint();
        fillPaint.Color = fillColor.Value;
        fillPaint.Style = SKPaintStyle.Fill;
        fillPaint.IsAntialias = true;

        canvas.DrawOval(rect, fillPaint);
    }

    private void RenderEllipseOutline(SKCanvas canvas, IShape shape, SKRect rect)
    {
        var outlineColor = this.GetShapeOutlineColor(shape);
        var strokeWidth = GetShapeOutlineWidth(shape);

        if (outlineColor is null || strokeWidth <= 0)
        {
            return;
        }

        using var outlinePaint = new SKPaint();
        outlinePaint.Color = outlineColor.Value;
        outlinePaint.Style = SKPaintStyle.Stroke;
        outlinePaint.StrokeWidth = (float)strokeWidth;
        outlinePaint.IsAntialias = true;

        canvas.DrawOval(rect, outlinePaint);
    }

    private SKColor? GetShapeFillColor(IShape shape)
    {
        var shapeFill = shape.Fill;

        // Check for explicit solid fill first
        if (shapeFill is { Type: FillType.Solid, Color: not null })
        {
            return new Color(shapeFill.Color).AsSkColor();
        }

        // Check for style-based fill (fillRef with scheme color)
        if (shapeFill is null || shapeFill.Type == FillType.NoFill)
        {
            var styleColor = this.GetStyleFillColor(shape);
            if (styleColor is not null)
            {
                return styleColor;
            }
        }

        return null;
    }

    private SKColor? GetShapeOutlineColor(IShape shape)
    {
        var shapeOutline = shape.Outline;

        // Check for explicit outline color first
        if (shapeOutline?.HexColor is not null)
        {
            return new Color(shapeOutline.HexColor).AsSkColor();
        }

        // Check for style-based outline (lnRef with scheme color)
        var styleColor = this.GetStyleOutlineColor(shape);
        if (styleColor is not null)
        {
            return styleColor;
        }

        return null;
    }

    private SKColor? GetStyleOutlineColor(IShape shape)
    {
        if (shape.SDKOpenXmlElement is not P.Shape { ShapeStyle.LineReference: { } lineRef })
        {
            return null;
        }

        var schemeColor = lineRef.GetFirstChild<A.SchemeColor>();
        if (schemeColor is null)
        {
            return null;
        }

        var schemeColorValue = schemeColor.Val?.InnerText;
        if (schemeColorValue is null)
        {
            return null;
        }

        var hexColor = this.ResolveSchemeColor(schemeColorValue);
        if (hexColor is null)
        {
            return null;
        }

        var baseColor = new Color(hexColor).AsSkColor();
        var shadeValue = schemeColor.GetFirstChild<A.Shade>()?.Val?.Value;

        return shadeValue is null
            ? baseColor
            : ApplyShade(baseColor, shadeValue.Value);
    }

    private SKColor? GetStyleFillColor(IShape shape)
    {
        var pShape = shape.SDKOpenXmlElement as P.Shape;
        if (pShape is null)
        {
            return null;
        }

        var style = pShape.ShapeStyle;
        var fillRef = style?.FillReference;
        if (fillRef is null)
        {
            return null;
        }

        var schemeColor = fillRef.GetFirstChild<A.SchemeColor>();
        if (schemeColor?.Val is null)
        {
            return null;
        }

        var schemeColorValue = schemeColor.Val?.InnerText;
        if (schemeColorValue is null)
        {
            return null;
        }

        var hexColor = this.ResolveSchemeColor(schemeColorValue);

        return hexColor is not null ? new Color(hexColor).AsSkColor() : null;
    }

    private string? ResolveSchemeColor(string schemeColorName)
    {
        var colorScheme = this.GetColorScheme();
        if (colorScheme is null)
        {
            return null;
        }

        if (!SchemeColorSelectors.TryGetValue(schemeColorName, out var selector))
        {
            return null;
        }

        var colorElement = selector(colorScheme);

        return colorElement is null ? null : GetHexFromColorElement(colorElement);
    }
    
    private A.ColorScheme? GetColorScheme() =>
        slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme.ThemeElements?.ColorScheme;
}
using System.Collections.Generic;
using System.IO;
using ImageMagick;
using ShapeCrawler.Shapes;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a shape collection.
/// </summary>
public interface ISlideShapeCollection : IShapeCollection
{
    /// <summary>
    ///     Adds the specified shape.
    /// </summary>
    void Add(IShape addingShape);

    /// <summary>
    ///     Adds a new audio shape.
    /// </summary>
    void AddAudio(int x, int y, Stream audio);

    /// <summary>
    ///     Adds a new audio shape.
    /// </summary>
    /// <param name="x">X coordinate in pixels.</param>
    /// <param name="y">Y coordinate in pixels.</param>
    /// <param name="audio">Audio stream.</param>
    /// <param name="type">Audio type.</param>
    void AddAudio(int x, int y, Stream audio, AudioType type);

    /// <summary>
    ///     Adds a new video from stream.
    /// </summary>
    /// <param name="x">X coordinate in pixels.</param>
    /// <param name="y">Y coordinate in pixels.</param>
    /// <param name="stream">Video stream data.</param>
    void AddVideo(int x, int y, Stream stream);

    /// <summary>
    ///     Adds a new shape.
    /// </summary>
    void AddShape(int x, int y, int width, int height, Geometry geometry = Geometry.Rectangle);
    
    /// <summary>
    ///     Adds a new shape.
    /// </summary>
    /// <param name="x">X coordinate in points.</param>
    /// <param name="y">Y coordinate in points.</param>
    /// <param name="width">Width in points.</param>
    /// <param name="height">Height in points.</param>
    /// <param name="geometry">Geometry form.</param>
    /// <param name="text">Text content.</param>
    void AddShape(int x, int y, int width, int height, Geometry geometry, string text);

    /// <summary>
    ///     Adds a line from XML.
    /// </summary>
    /// <param name="xml">Content of p:cxnSp Open XML element.</param>
    void AddLine(string xml);

    /// <summary>
    ///     Adds a new line.
    /// </summary>
    void AddLine(int startPointX, int startPointY, int endPointX, int endPointY);

    /// <summary>
    ///     Adds a new table.
    /// </summary>
    void AddTable(int x, int y, int columnsCount, int rowsCount);

    /// <summary>
    ///     Adds a new table with a custom style.
    /// </summary>
    void AddTable(int x, int y, int columnsCount, int rowsCount, ITableStyle style);

    /// <summary>
    ///     Adds picture.
    /// </summary>
    void AddPicture(Stream imageStream, MagickFormat format = MagickFormat.Unknown);

    /// <summary>
    ///     Adds Pie Chart.
    /// </summary>
    void AddPieChart(int x, int y, int width, int height, Dictionary<string, double> categoryValues, string seriesName);
    
    /// <summary>
    ///     Adds Bar Chart with specified parameters.
    /// </summary>
    void AddBarChart(int x, int y, int width, int height, Dictionary<string, double> categoryValues, string seriesName);
    
    /// <summary>
    ///     Adds Scatter Chart.
    /// </summary>
    /// <param name="x">X coordinate in points.</param>
    /// <param name="y">Y coordinate in points.</param>
    /// <param name="width">Width in point.</param>
    /// <param name="height">Height in points.</param>
    /// <param name="pointValues">Dictionary of x and y coordinate values for each point.</param>
    /// <param name="seriesName">Name of the data series.</param>
    void AddScatterChart(int x, int y, int width, int height, Dictionary<double, double> pointValues, string seriesName);

    /// <summary>
    ///     Adds a Stacked Column Chart.
    /// </summary>
    /// <param name="x">X coordinate in points.</param>
    /// <param name="y">Y coordinate in points.</param>
    /// <param name="width">Width in point.</param>
    /// <param name="height">Height in points.</param>
    /// <param name="categoryValues">Dictionary mapping categories to a list of values for each series.</param>
    /// <param name="seriesNames">List of series names in the same order as the values in categoryValues.</param>
    void AddStackedColumnChart(int x, int y, int width, int height, IDictionary<string, IList<double>> categoryValues, IList<string> seriesNames);

    /// <summary>
    ///     Adds shape with SmartArt graphic content.
    /// </summary>
    IShape AddSmartArt(int x, int y, int width, int height, SmartArtType smartArtType);

    /// <summary>
    ///     Groups the specified shapes.
    /// </summary>
    IShape Group(IShape[] groupingShapes);
}
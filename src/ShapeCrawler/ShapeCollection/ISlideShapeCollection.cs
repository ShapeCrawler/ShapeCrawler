using System.Collections.Generic;
using System.IO;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a shape collection.
/// </summary>
public interface ISlideShapeCollection : IShapeCollection
{
    /// <summary>
    ///     Adds the specified shape.
    /// </summary>
    void Add(IShape shape);

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
    ///     Adds a new Rectangle shape.
    /// </summary>
    void AddShape(int x, int y, int width, int height, Geometry geometry = Geometry.Rectangle);

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
    ///     Removes specified shape.
    /// </summary>
    void Remove(IShape shape);

    /// <summary>
    ///     Adds picture.
    /// </summary>
    void AddPicture(Stream image);

    /// <summary>
    ///     Adds Pie Chart.
    /// </summary>
    void AddPieChart(int x, int y, int width, int height, Dictionary<string, double> categoryValues, string seriesName);
}
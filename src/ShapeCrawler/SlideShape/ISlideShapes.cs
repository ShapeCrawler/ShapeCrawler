using System.IO;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a shape collection.
/// </summary>
public interface ISlideShapes : IReadOnlyShapeCollection
{
    /// <summary>
    ///     Adds a new shape from other shape.
    /// </summary>
    void Add(IShape shape);

    /// <summary>
    ///     Adds a new audio from stream.
    /// </summary>
    /// <param name="xPixel">The X coordinate for the left side of the shape.</param>
    /// <param name="yPixels">The Y coordinate for the left side of the shape.</param>
    /// <param name="mp3Stream">Audio stream data.</param>
    void AddAudio(int xPixel, int yPixels, Stream mp3Stream);

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
    void AddRectangle(int x, int y, int w, int h);

    /// <summary>
    ///     Adds a new Rounded Rectangle shape. 
    /// </summary>
    void AddRoundedRectangle(int x, int y, int w, int h);

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
    ///     Removes specified shape.
    /// </summary>
    void Remove(IShape shape);

    /// <summary>
    ///     Adds picture.
    /// </summary>
    void AddPicture(Stream imageStream);

    /// <summary>
    ///     Adds Bar chart.
    /// </summary>
    void AddBarChart(BarChartType barChartType);
}
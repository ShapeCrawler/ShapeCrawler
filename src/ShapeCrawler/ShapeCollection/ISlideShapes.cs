using System.IO;
using ShapeCrawler.Shared;
using Svg;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a shape collection.
/// </summary>
public interface ISlideShapes : IShapes
{
    /// <summary>
    ///     Adds a new shape from other shape.
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
    void AddRectangle(int x, int y, int width, int height);

    /// <summary>
    ///     Adds a new Rectangle: Rounded Corners. 
    /// </summary>
    void AddRoundedRectangle(int x, int y, int width, int height);

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
    ///     Adds an SVG picture
    /// </summary>
    /// <remarks>
    ///     Note that an SVG picture is inserted into the document as BOTH a
    ///     vector image AND a rasterized image. The rasterized image is what
    ///     the user sees. The vector image is used by the application to re-
    ///     rasterize at a different size later.
    /// </remarks>
    /// <param name="image">The image to be loaded</param>
    /// <param name="rasterWidth">Width of rasterized image in pixels.</param>
    /// <param name="rasterHeight">Height of rasterized image in pixels.</param>
    void AddPictureSvg(SvgDocument image, int rasterWidth, int rasterHeight);

    /// <summary>
    ///     Adds Bar chart.
    /// </summary>
    void AddBarChart(BarChartType barChartType);
}
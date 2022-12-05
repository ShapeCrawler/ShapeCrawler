using System.Collections.Generic;
using System.IO;
using ShapeCrawler.Media;
using ShapeCrawler.Shapes;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a shape collection.
/// </summary>
public interface IShapeCollection : IEnumerable<IShape>
{
    /// <summary>
    ///     Gets the number of series items in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    ///     Gets shape at the specified index.
    /// </summary>
    IShape this[int index] { get; }

    /// <summary>
    ///     Gets shape by identifier.
    /// </summary>
    /// <typeparam name="T">The type of shape.</typeparam>
    T GetById<T>(int shapeId)
        where T : IShape;

    /// <summary>
    ///     Get shape by name.
    /// </summary>
    /// <typeparam name="T">The type of shape.</typeparam>
    T GetByName<T>(string shapeName)
        where T : IShape;

    /// <summary>
    ///     Create a new audio shape from stream and adds it to the end of the collection.
    /// </summary>
    /// <param name="xPixel">The X coordinate for the left side of the shape.</param>
    /// <param name="yPixels">The Y coordinate for the left side of the shape.</param>
    /// <param name="mp3Stream">Audio stream data.</param>
    IAudioShape AddAudio(int xPixel, int yPixels, Stream mp3Stream);

    /// <summary>
    ///     Create a new video shape from stream and adds it to the end of the collection.
    /// </summary>
    /// <param name="x">X coordinate in pixels.</param>
    /// <param name="y">Y coordinate in pixels.</param>
    /// <param name="stream">Video stream data.</param>
    IVideoShape AddVideo(int x, int y, Stream stream);

    /// <summary>
    ///     Creates a new AutoShape.
    /// </summary>
    /// <param name="geometry">Geometry form.</param>
    /// <param name="x">X coordinate in pixels.</param>
    /// <param name="y">Y coordinate in pixels.</param>
    /// <param name="width">Width in pixels.</param>
    /// <param name="height">Height in pixels.</param>
    IAutoShape AddAutoShape(SCGeometry geometry, int x, int y, int width, int height);
}
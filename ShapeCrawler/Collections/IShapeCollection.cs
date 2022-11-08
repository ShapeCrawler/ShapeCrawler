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
    IAudioShape AddNewAudio(int xPixel, int yPixels, Stream mp3Stream);

    /// <summary>
    ///     Create a new video shape from stream and adds it to the end of the collection.
    /// </summary>
    /// <param name="xPixel">The X coordinate for the left side of the shape.</param>
    /// <param name="yPixels">The Y coordinate for the left side of the shape.</param>
    /// <param name="videoStream">Video stream data.</param>
    IVideoShape AddNewVideo(int xPixel, int yPixels, Stream videoStream);
}
using System.IO;

// ReSharper disable CheckNamespace
#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents an image.
/// </summary>
public interface IImage
{
    /// <summary>
    ///     Gets MIME type.
    /// </summary>
    string Mime { get; }

    /// <summary>
    ///     Gets file name of internal resource.
    /// </summary>
    string Name { get; }

    /// <summary>
    ///     Gets binary content.
    /// </summary>
    byte[] AsByteArray();

    /// <summary>
    ///     Sets image with stream.
    /// </summary>
    void Update(Stream stream);
}
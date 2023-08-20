using System.IO;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents an image.
/// </summary>
public interface IImage
{
    /// <summary>
    ///     Gets MIME type.
    /// </summary>
    string MIME { get; }

    /// <summary>
    ///     Gets file name of internal resource.
    /// </summary>
    string Name { get; }

    /// <summary>
    ///     Gets binary content.
    /// </summary>
    byte[] BinaryData();

    /// <summary>
    ///     Sets image with stream.
    /// </summary>
    void Update(Stream stream);

    /// <summary>
    ///     Sets image with byte array.
    /// </summary>
    void Update(byte[] bytes);

    /// <summary>
    ///     Sets image by specified file path.
    /// </summary>
    void Update(string filePath);
}
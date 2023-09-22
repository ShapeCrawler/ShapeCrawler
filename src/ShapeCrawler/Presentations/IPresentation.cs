using System.IO;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a presentation document.
/// </summary>
public interface IPresentation : IPresentationProperties
{
    /// <summary>
    ///     Saves presentation in specified file path.
    /// </summary>
    void SaveAs(string path);

    /// <summary>
    ///     Saves presentation in specified stream.
    /// </summary>
    void SaveAs(Stream stream);
}
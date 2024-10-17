using System.IO;

// ReSharper disable once CheckNamespace
#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a presentation document.
/// </summary>
public interface IPresentation : IPresentationProperties
{
    /// <summary>
    ///     Saves presentation in the specified file path.
    /// </summary>
    void SaveAs(string path);

    /// <summary>
    ///     Saves presentation in specified stream.
    /// </summary>
    void SaveAs(Stream stream);
}
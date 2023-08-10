using System.IO;
using DocumentFormat.OpenXml.Packaging;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

internal interface ISavePresentation : IPresentation
{
    /// <summary>
    ///     Saves presentation in specified file path.
    /// </summary>
    void Save(string path);

    /// <summary>
    ///     Saves presentation in specified stream.
    /// </summary>
    void Save(Stream stream);
}
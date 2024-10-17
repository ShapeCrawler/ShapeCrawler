using System.IO;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a presentation.
/// </summary>
internal interface IValidateable : IPresentationProperties
{
    /// <summary>
    ///     Validates presentation.
    /// </summary>
    void Validate();
    
    /// <summary>
    ///     Saves presentation to the specified path.
    /// </summary>
    void CopyTo(string path);
    
    /// <summary>
    ///     Saves presentation to the specified stream.
    /// </summary>
    void CopyTo(Stream stream);
}
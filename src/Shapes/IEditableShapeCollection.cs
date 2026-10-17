using System.IO;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a collection of shapes that can be edited.
/// </summary>
public interface IEditableShapeCollection : IShapeCollection
{
    /// <summary>
    ///     Adds a picture from the specified image stream.
    /// </summary>
    /// <param name="imageStream">Image stream.</param>
    void AddPicture(Stream imageStream);
}
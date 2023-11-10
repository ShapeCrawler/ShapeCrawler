namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents path information of the shape.
/// </summary>
public interface IPath
{
    /// <summary>
    ///     Gets the shape's XPath in the slide.
    /// </summary>
    public string SDKXPath { get; }
}
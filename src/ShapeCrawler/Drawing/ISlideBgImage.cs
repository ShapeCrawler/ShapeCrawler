namespace ShapeCrawler;

/// <summary>
///     Represents a slide background image.
/// </summary>
public interface ISlideBgImage : IImage
{
    /// <summary>
    ///     Presents whether the background image is presented.
    /// </summary>
    /// <returns></returns>
    bool Present();
}
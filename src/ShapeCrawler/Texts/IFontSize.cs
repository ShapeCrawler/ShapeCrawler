namespace ShapeCrawler.Texts;

/// <summary>
///     Represents a font size.
/// </summary>
internal interface IFontSize
{
    /// <summary>
    ///     Returns font size in points.
    /// </summary>
    int Size();
    
    /// <summary>
    ///     Updates font size in points.
    /// </summary>
    void Update(int points);
}
namespace ShapeCrawler.Presentations;

/// <summary>
///    Represents ShapeCrawler settings.
/// </summary>
public static class SCSettings
{
    /// <summary>
    ///    Gets or sets time provider.
    /// </summary>
    public static ITimeProvider TimeProvider { get; set; } = new SystemTimeProvider();
}
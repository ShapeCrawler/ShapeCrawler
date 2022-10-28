// ReSharper disable once CheckNamespace

namespace ShapeCrawler;
#if DEBUG

/// <summary>
///     ShapeCrawlers settings.
/// </summary>
public static class SCSettings
{
    /// <summary>
    ///     Gets or sets a value indicating whether ShapeCrawler can collect statistic. The default value is <c>true</c>.
    /// </summary>
    public static bool CanCollectStatistic { get; set; } = true;
}

#endif
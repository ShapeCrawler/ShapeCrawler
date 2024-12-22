namespace ShapeCrawler.Shared;

internal static class ShapeCrawlerInternal
{
    internal static ITimeProvider TimeProvider { get; set; } = new SystemTimeProvider();
}
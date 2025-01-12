namespace ShapeCrawler.Shared;

internal static class SCSettings
{
    internal static ITimeProvider TimeProvider { get; set; } = new SystemTimeProvider();
}
namespace ShapeCrawler.Presentations;

public static class SCSettings
{
    public static ITimeProvider TimeProvider { get; set; } = new SystemTimeProvider();
}
namespace ShapeCrawler.Charts;

internal readonly record struct ChartLayout(
    float CenterX,
    float CenterY,
    float Radius,
    float AvailableWidth);
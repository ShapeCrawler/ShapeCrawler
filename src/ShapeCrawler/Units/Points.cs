namespace ShapeCrawler.Units;

internal readonly ref struct Points(float points)
{
    internal long AsEmus() => (long)(points * 12700);

    internal float AsPixels() => points * 96 / 72;
}
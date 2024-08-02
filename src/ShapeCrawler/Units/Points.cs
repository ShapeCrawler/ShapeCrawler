namespace ShapeCrawler.Units;

internal readonly ref struct Points(float points)
{
    internal long AsEmus() => (long)(points * 12700);
}
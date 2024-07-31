namespace ShapeCrawler.Units;

internal readonly ref struct Points(float points)
{
    public long AsEmus()
    {
        return (long)(points * 12700);
    }
}
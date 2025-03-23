namespace ShapeCrawler.Units;

internal readonly ref struct Emus(long emus)
{
    internal decimal AsPoints() => emus / 12700m;
}
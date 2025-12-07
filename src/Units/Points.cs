namespace ShapeCrawler.Units;

internal readonly ref struct Points(decimal points)
{
    internal long AsEmus() => (long)(points * 12700);
    
    internal int AsHundredPoints() => (int)(points * 100);

    internal decimal AsPixels()
    {
        const decimal pointsPerInch = 72m;
        const decimal pixelsPerInch = 96m;
        
        return points * pixelsPerInch / pointsPerInch;
    }
}
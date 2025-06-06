﻿namespace ShapeCrawler.Units;

internal readonly ref struct Points(decimal points)
{
    internal long AsEmus() => (long)(points * 12700);
    
    internal int AsHundredPoints() => (int)(points * 100);
}
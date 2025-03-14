using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Units;

namespace ShapeCrawler.Tables;

internal readonly ref struct TopBottomMargin(Int32Value? emus)
{
    const float DefaultTopAndBottomMargin = 3.69f; // ~0.13 cm

    public float Value => emus is null ? DefaultTopAndBottomMargin : new Emus(emus).AsPoints();
}
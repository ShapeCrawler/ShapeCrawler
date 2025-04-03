using DocumentFormat.OpenXml;
using ShapeCrawler.Units;

namespace ShapeCrawler.Tables;

internal readonly ref struct TopBottomMargin(Int32Value? emus)
{
    private const decimal DefaultTopAndBottomMargin = 3.69m; // ~0.13 cm

    public decimal Value => emus is null ? DefaultTopAndBottomMargin : new Emus(emus).AsPoints();
}
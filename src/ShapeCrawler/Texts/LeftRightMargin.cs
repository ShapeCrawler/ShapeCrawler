using DocumentFormat.OpenXml;
using ShapeCrawler.Units;

namespace ShapeCrawler.Texts;

internal readonly ref struct LeftRightMargin(Int32Value? emus)
{
    const decimal DefaultLeftAndRightMarginPoints = 7.09m; // ~0.25 cm

    public decimal Value => emus is null ? DefaultLeftAndRightMarginPoints : new Emus(emus).AsPoints();
}
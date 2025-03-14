using DocumentFormat.OpenXml;
using ShapeCrawler.Units;

namespace ShapeCrawler.Texts;

internal readonly ref struct LeftRightMargin(Int32Value? emus)
{
    const float DefaultLeftAndRightMarginPoints = 7.09f; // ~0.25 cm

    public float Value => emus is null ? DefaultLeftAndRightMarginPoints : new Emus(emus).AsPoints();
}
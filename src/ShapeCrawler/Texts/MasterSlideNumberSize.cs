using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal record MasterSlideNumberSize : IFontSize
{
    private readonly A.DefaultRunProperties aDefaultRunProperties;
    private const int HalfPointsInPoint = 100;

    internal MasterSlideNumberSize(A.DefaultRunProperties aDefaultRunProperties)
    {
        this.aDefaultRunProperties = aDefaultRunProperties;
    }

    public int Size()
    {
        var halfPoints = this.aDefaultRunProperties.FontSize!.Value;
        return halfPoints / HalfPointsInPoint;
    }

    public void Update(int points)
    {
        var halfPoints = points * HalfPointsInPoint;
        this.aDefaultRunProperties.FontSize!.Value = halfPoints;
    }
}
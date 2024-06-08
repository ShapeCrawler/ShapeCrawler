using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal record MasterSlideNumberSize : IFontSize
{
    private readonly A.DefaultRunProperties aDefaultRunProperties;
    private const decimal HalfPointsInPoint = 100m;

    internal MasterSlideNumberSize(A.DefaultRunProperties aDefaultRunProperties)
    {
        this.aDefaultRunProperties = aDefaultRunProperties;
    }

    public decimal Size()
    {
        var halfPoints = this.aDefaultRunProperties.FontSize!.Value;
        return halfPoints / HalfPointsInPoint;
    }

    public void Update(decimal points)
    {
        var halfPoints = points * HalfPointsInPoint;
        this.aDefaultRunProperties.FontSize!.Value = (int)halfPoints;
    }
}
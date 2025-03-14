using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal record MasterSlideNumberSize : IFontSize
{
    private readonly A.DefaultRunProperties aDefaultRunProperties;

    internal MasterSlideNumberSize(A.DefaultRunProperties aDefaultRunProperties)
    {
        this.aDefaultRunProperties = aDefaultRunProperties;
    }

    public float Size()
    {
        var hundredsOfPoint = this.aDefaultRunProperties.FontSize!.Value;
        
        return hundredsOfPoint / 100f;
    }

    public void Update(float points)
    {
        var halfPoints = points * 100;
        this.aDefaultRunProperties.FontSize!.Value = (int)halfPoints;
    }
}
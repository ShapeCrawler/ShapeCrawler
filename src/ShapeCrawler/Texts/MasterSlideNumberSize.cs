using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal class MasterSlideNumberSize(A.DefaultRunProperties aDefaultRunProperties): IFontSize
{
    public decimal Size
    {
        get
        {
            var hundredPoints = aDefaultRunProperties.FontSize!.Value;
        
            return hundredPoints / 100m;
        }

        set
        {
            var hundredPoints = value * 100;
            aDefaultRunProperties.FontSize!.Value = (int)hundredPoints;
        }
    }
}
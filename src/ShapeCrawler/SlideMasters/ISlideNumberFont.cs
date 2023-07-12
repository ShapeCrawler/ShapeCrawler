using ShapeCrawler.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler;

public interface ISlideNumberFont
{
    SCColor Color { get; set; }
}

internal class SCSlideNumberFont : ISlideNumberFont
{
    private readonly A.DefaultRunProperties aDefaultRunProperties;

    internal SCSlideNumberFont(A.DefaultRunProperties aDefaultRunProperties)
    {
        this.aDefaultRunProperties = aDefaultRunProperties;
    }

    public SCColor Color
    {
        get => this.ParseColor();
        set => this.UpdateColor(value);
    }

    private void UpdateColor(SCColor color)
    {
        var solidFill = this.aDefaultRunProperties.GetFirstChild<A.SolidFill>()!;
        solidFill.RemoveAllChildren();
        
        var rgbColorModelHex = new A.RgbColorModelHex
        {
            Val = color.ToString()
        };
        solidFill.AppendChild(rgbColorModelHex);
    }

    private SCColor ParseColor()
    {
        var hex = this.aDefaultRunProperties.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value!;

        return SCColor.FromHex(hex);
    }
}
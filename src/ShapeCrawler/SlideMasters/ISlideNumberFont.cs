using ShapeCrawler.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a slide number font.
/// </summary>
public interface ISlideNumberFont
{
    /// <summary>
    ///     Gets or sets color.
    /// </summary>
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
        var solidFill = this.aDefaultRunProperties.GetFirstChild<A.SolidFill>();
        solidFill?.Remove();

        var rgbColorModelHex = new A.RgbColorModelHex { Val = color.ToString() };
        solidFill = new A.SolidFill(rgbColorModelHex);
        
        this.aDefaultRunProperties.Append(solidFill);
    }

    private SCColor ParseColor()
    {
        var hex = this.aDefaultRunProperties.GetFirstChild<A.SolidFill>() !.RgbColorModelHex!.Val!.Value!;

        return SCColor.FromHex(hex);
    }
}
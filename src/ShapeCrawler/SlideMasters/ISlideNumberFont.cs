using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a slide number font.
/// </summary>
public interface ISlideNumberFont : IFont
{
    /// <summary>
    ///     Gets or sets color.
    /// </summary>
    SCColor Color { get; set; }
}

internal sealed class SCSlideNumberFont : ISlideNumberFont
{
    private readonly A.DefaultRunProperties aDefaultRunProperties;
    private readonly List<ITextPortionFont> portionFonts;

    internal SCSlideNumberFont(A.DefaultRunProperties aDefaultRunProperties, List<ITextPortionFont> portionFonts)
    {
        this.aDefaultRunProperties = aDefaultRunProperties;
        this.portionFonts = portionFonts;
    }

    public SCColor Color
    {
        get => this.ParseColor();
        set => this.UpdateColor(value);
    }

    public int Size
    {
        get => this.ParseSize();
        set => this.UpdateSize(value);
    }

    private void UpdateSize(int points)
    {
        this.portionFonts.ForEach(pf => pf.Size = points);
    }

    private int ParseSize()
    {
        return this.portionFonts.First().Size;
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
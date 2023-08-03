using System.Collections.Generic;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a slide number.
/// </summary>
public interface IMasterSlideNumber : IShapeLocation
{
    /// <summary>
    ///     Gets font.
    /// </summary>
    ISlideNumberFont Font { get; }
}

internal sealed class SCMasterSlideNumber : IMasterSlideNumber
{
    private readonly ShapeLocation shapeLocation;

    internal SCMasterSlideNumber(P.Shape pSldNum)
    {
        var aDefaultRunProperties =
            pSldNum.TextBody!.ListStyle!.Level1ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>() !;
        this.Font = new SCSlideNumberFont(aDefaultRunProperties);
        this.shapeLocation = new ShapeLocation(pSldNum.ShapeProperties!.Transform2D!);
    }

    public ISlideNumberFont Font { get; }

    public int X
    {
        get => this.shapeLocation.ParseX();
        set => this.shapeLocation.UpdateX(value);
    }

    public int Y
    {
        get => this.shapeLocation.ParseY();
        set => this.shapeLocation.UpdateY(value);
    }
}
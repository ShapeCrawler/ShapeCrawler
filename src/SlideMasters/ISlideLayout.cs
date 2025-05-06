using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;

// ReSharper disable CheckNamespace
#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a Slide Layout.
/// </summary>
public interface ISlideLayout
{
    /// <summary>
    ///     Gets layout name.
    /// </summary>
    string Name { get; }

    /// <summary>
    ///     Gets layout shape collection.
    /// </summary>
    IShapeCollection Shapes { get; }

    /// <summary>
    ///     Gets slide master.
    /// </summary>
    ISlideMaster SlideMaster { get; }

    /// <summary>
    ///     Gets layout number.
    /// </summary>
    int Number { get; }
}

internal sealed class SlideLayout : ISlideLayout
{
    private readonly SlideLayoutPart slideLayoutPart;

    internal SlideLayout(SlideLayoutPart sdkLayoutPart)
        : this(sdkLayoutPart, new SlideMaster(sdkLayoutPart.SlideMasterPart!))
    {
    }

    private SlideLayout(SlideLayoutPart sdkLayoutPart, ISlideMaster slideMaster)
    {
        this.slideLayoutPart = sdkLayoutPart;
        this.SlideMaster = slideMaster;
        this.Shapes = new ShapeCollection(this.slideLayoutPart);
    }

    public string Name => this.slideLayoutPart.SlideLayout.CommonSlideData!.Name!.Value!;

    public IShapeCollection Shapes { get; }

    public ISlideMaster SlideMaster { get; }

    public int Number => int.Parse(Regex.Match(this.slideLayoutPart.Uri.ToString(), @"\d+").Value);

    internal SlideLayoutPart SdkSlideLayoutPart() => this.slideLayoutPart;
}
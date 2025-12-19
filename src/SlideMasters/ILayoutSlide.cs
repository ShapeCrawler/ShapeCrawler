using System;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;

// ReSharper disable CheckNamespace
#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a Slide Layout in PowerPoint UI.
/// </summary>
public interface ILayoutSlide
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
    IMasterSlide MasterSlide { get; }

    /// <summary>
    ///     Gets layout number.
    /// </summary>
    int Number { get; }

    /// <summary>
    ///     Gets layout background.
    /// </summary>
    ILayoutSlideBackground Background { get; }
}

internal sealed class LayoutSlide : ILayoutSlide
{
    private readonly LayoutSlideBackground background;

    internal LayoutSlide(SlideLayoutPart slideLayoutPart)
    {
        this.SlideLayoutPart = slideLayoutPart;
        this.Shapes = new ShapeCollection(slideLayoutPart);
        this.background = new LayoutSlideBackground(slideLayoutPart);
    }
    
    public string Name => this.SlideLayoutPart.SlideLayout.CommonSlideData!.Name!.Value!;

    public IShapeCollection Shapes { get; }

    public IMasterSlide MasterSlide => new MasterSlide(this.SlideLayoutPart.SlideMasterPart!);

    public int Number
    {
        get
        {
            var match = Regex.Match(this.SlideLayoutPart.Uri.ToString(), @"\d+", RegexOptions.None, TimeSpan.FromSeconds(1));
            return int.Parse(match.Value);
        }
    }

    public ILayoutSlideBackground Background => this.background;
    
    internal SlideLayoutPart SlideLayoutPart { get; }
}
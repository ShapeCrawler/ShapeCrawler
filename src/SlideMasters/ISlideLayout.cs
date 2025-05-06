using System;
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
    internal SlideLayout(SlideLayoutPart slideLayoutPart)
    {
        this.SlideLayoutPart = slideLayoutPart;
        this.Shapes = new ShapeCollection(slideLayoutPart);
    }
    
    public string Name => this.SlideLayoutPart.SlideLayout.CommonSlideData!.Name!.Value!;

    public IShapeCollection Shapes { get; }

    public ISlideMaster SlideMaster => new SlideMaster(this.SlideLayoutPart.SlideMasterPart!);

    public int Number
    {
        get
        {
            var match = Regex.Match(this.SlideLayoutPart.Uri.ToString(), @"\d+", RegexOptions.None, TimeSpan.FromSeconds(1));
            return int.Parse(match.Value);
        }
    }
    
    internal SlideLayoutPart SlideLayoutPart { get; }
}
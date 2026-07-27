using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a PowerPoint Slide Layout.
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

    /// <summary>
    ///     Moves the layout to the specified slide master.
    /// </summary>
    /// <param name="targetMaster">Slide master that will own the layout.</param>
    void MoveTo(IMasterSlide targetMaster);
}

internal sealed class LayoutSlide : ILayoutSlide
{
    private readonly LayoutSlideBackground background;
    private readonly SlideLayoutPart slideLayoutPart;

    internal LayoutSlide(SlideLayoutPart slideLayoutPart)
    {
        this.slideLayoutPart = slideLayoutPart;
        this.Shapes = new ShapeCollection(slideLayoutPart);
        this.background = new LayoutSlideBackground(slideLayoutPart);
    }

    public string Name => this.slideLayoutPart.SlideLayout!.CommonSlideData!.Name!.Value!;

    public IShapeCollection Shapes { get; }

    public IMasterSlide MasterSlide => new MasterSlide(this.slideLayoutPart.SlideMasterPart!);

    public int Number
    {
        get
        {
            var match = Regex.Match(this.slideLayoutPart.Uri.ToString(), @"\d+", RegexOptions.None, TimeSpan.FromSeconds(1));
            return int.Parse(match.Value);
        }
    }

    public ILayoutSlideBackground Background => this.background;

    /// <inheritdoc />
    public void MoveTo(IMasterSlide targetMaster)
    {
        var targetSlideMasterPart = ((MasterSlide)targetMaster).InternalSlideMasterPart();
        var sourceSlideMasterPart = this.slideLayoutPart.SlideMasterPart!;
        if (sourceSlideMasterPart == targetSlideMasterPart)
        {
            return;
        }

        var sourceRelationshipId = sourceSlideMasterPart.GetIdOfPart(this.slideLayoutPart);
        var pSlideLayoutId = sourceSlideMasterPart.SlideMaster!.SlideLayoutIdList!
            .Elements<P.SlideLayoutId>()
            .Single(layoutId => layoutId.RelationshipId == sourceRelationshipId);

        targetSlideMasterPart.AddPart(this.slideLayoutPart);
        var targetRelationshipId = targetSlideMasterPart.GetIdOfPart(this.slideLayoutPart);
        pSlideLayoutId.Remove();
        pSlideLayoutId.RelationshipId = targetRelationshipId;
        targetSlideMasterPart.SlideMaster!.SlideLayoutIdList!.Append(pSlideLayoutId);
        sourceSlideMasterPart.DeletePart(sourceRelationshipId);
    }

    /// <summary>
    ///     Gets the underlying Open XML slide layout part.
    /// </summary>
    /// <returns>Open XML slide layout part.</returns>
    internal SlideLayoutPart InternalSlideLayoutPart() => this.slideLayoutPart;
}
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.SlideMasters;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a collections of Slide Masters.
/// </summary>
public interface ISlideMasterCollection
{
    /// <summary>
    ///     Gets the number of series items in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    ///     Gets the element at the specified index.
    /// </summary>
    ISlideMaster this[int index] { get; }

    /// <summary>
    ///     Gets the generic enumerator that iterates through the collection.
    /// </summary>
    IEnumerator<ISlideMaster> GetEnumerator();
}

internal sealed class SlideMasterCollection : ISlideMasterCollection
{
    private readonly List<ISlideMaster> slideMasters;

    private SlideMasterCollection(SCPresentation presentation, List<ISlideMaster> slideMasters)
    {
        this.Presentation = presentation;
        this.slideMasters = slideMasters;
    }

    public int Count => this.slideMasters.Count;

    internal SCPresentation Presentation { get; }

    public ISlideMaster this[int index] => this.slideMasters[index];

    public IEnumerator<ISlideMaster> GetEnumerator()
    {
        return this.slideMasters.GetEnumerator();
    }

    internal static SlideMasterCollection Create(SCPresentation presentation)
    {
        var masterParts = presentation.SDKPresentationInternal.PresentationPart!.SlideMasterParts;
        var slideMasters = new List<ISlideMaster>(masterParts.Count());
        var number = 1;
        foreach (var slideMasterPart in masterParts)
        {
            slideMasters.Add(new SCSlideMaster(presentation, slideMasterPart.SlideMaster, number++));
        }

        return new SlideMasterCollection(presentation, slideMasters);
    }

    internal SCSlideLayout GetSlideLayoutBySlide(SCSlide slide)
    {
        SlideLayoutPart inputSlideLayoutPart = slide.SDKSlidePart.SlideLayoutPart!;
        IEnumerable<SCSlideLayout> allLayouts = this.slideMasters.SelectMany(sm => sm.SlideLayouts).OfType<SCSlideLayout>();

        return allLayouts.First(sl => sl.SlideLayoutPart.Uri == inputSlideLayoutPart.Uri);
    }
}
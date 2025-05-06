using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMasters;

/// <summary>
///     Represents a slide layout collection.
/// </summary>
public interface ISlideLayoutCollection : IEnumerable<ISlideLayout>
{
    /// <summary>
    ///     Gets slide layout by index.
    /// </summary>
    ISlideLayout this[int index] { get; }
}

internal sealed class SlideLayoutCollection(SlideMasterPart slideMasterPart) : ISlideLayoutCollection
{
    public ISlideLayout this[int index] => this.Layouts().ElementAt(index);

    public IEnumerator<ISlideLayout> GetEnumerator() => this.Layouts().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    internal SlideLayout Layout(int number) => this.Layouts().First(l => l.Number == number);

    private IEnumerable<SlideLayout> Layouts()
    {
        var rIdList = slideMasterPart.SlideMaster.SlideLayoutIdList!.OfType<P.SlideLayoutId>()
            .Select(layoutId => layoutId.RelationshipId!);
        foreach (var rId in rIdList)
        {
            var layoutPart = (SlideLayoutPart)slideMasterPart.GetPartById(rId.Value!);

            yield return new SlideLayout(layoutPart);
        }
    }
}
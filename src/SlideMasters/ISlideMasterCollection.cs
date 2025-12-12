using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

// ReSharper disable PossibleMultipleEnumeration
#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents slide master collection.
/// </summary>
public interface ISlideMasterCollection : IEnumerable<IMasterSlide>
{
    /// <summary>
    ///     Gets slide master by index.
    /// </summary>
    IMasterSlide this[int index] { get; }
}

internal sealed class SlideMasterCollection(IEnumerable<SlideMasterPart> slideMasterParts) : ISlideMasterCollection
{
    public IMasterSlide this[int index] => this.SlideMasters().ElementAt(index);

    public IEnumerator<IMasterSlide> GetEnumerator() => this.SlideMasters().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    internal MasterSlide SlideMaster(int number) =>
        this.SlideMasters().First(slideMaster => slideMaster.Number == number);

    private IEnumerable<MasterSlide> SlideMasters()
    {
        foreach (var slideMaster in slideMasterParts)
        {
            yield return new MasterSlide(slideMaster);
        }
    }
}
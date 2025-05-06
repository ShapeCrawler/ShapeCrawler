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
public interface ISlideMasterCollection : IEnumerable<ISlideMaster>
{
    /// <summary>
    ///     Gets slide master by index.
    /// </summary>
    ISlideMaster this[int index] { get; }
}

internal sealed class SlideMasterCollection(IEnumerable<SlideMasterPart> slideMasterParts) : ISlideMasterCollection
{
    public ISlideMaster this[int index] => this.SlideMasters().ElementAt(index);

    public IEnumerator<ISlideMaster> GetEnumerator() => this.SlideMasters().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    internal SlideMaster SlideMaster(int number) =>
        this.SlideMasters().First(slideMaster => slideMaster.Number == number);

    private IEnumerable<SlideMaster> SlideMasters()
    {
        foreach (var slideMaster in slideMasterParts)
        {
            yield return new SlideMaster(slideMaster);
        }
    }
}
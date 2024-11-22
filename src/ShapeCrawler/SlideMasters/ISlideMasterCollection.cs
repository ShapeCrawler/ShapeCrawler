using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a collections of Slide Masters.
/// </summary>
public interface ISlideMasterCollection : IEnumerable<ISlideMaster>
{
    /// <summary>
    ///     Gets the number of series items in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    ///     Gets the element at the specified index.
    /// </summary>
    ISlideMaster this[int index] { get; }
}

internal sealed class SlideMasterCollection : ISlideMasterCollection
{
    private readonly List<ISlideMaster> slideMasters;

    internal SlideMasterCollection(IEnumerable<SlideMasterPart> sdkMasterParts)
    {
        this.slideMasters = new List<ISlideMaster>(sdkMasterParts.Count());
        foreach (var sdkMasterPart in sdkMasterParts)
        {
            this.slideMasters.Add(new SlideMaster(sdkMasterPart));
        }
    }
    
    public int Count => this.slideMasters.Count;

    public ISlideMaster this[int index] => this.slideMasters[index];

    public IEnumerator<ISlideMaster> GetEnumerator()
    {
        return this.slideMasters.GetEnumerator();
    }
    
    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
}
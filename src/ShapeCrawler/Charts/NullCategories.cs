using System.Collections;
using System.Collections.Generic;

namespace ShapeCrawler.Charts;

internal class NullCategories : IReadOnlyList<ICategory>
{
    private const string Error =
        $"Chart does not have categories. Use {nameof(IChart.HasCategories)} property to check if chart categories are available.";
    
    public int Count => throw new SCException(Error);
    
    public ICategory this[int index] => throw new SCException(Error);
    
    public IEnumerator<ICategory> GetEnumerator() => throw new SCException(Error);
    
    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();
}
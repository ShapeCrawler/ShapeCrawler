using System.Collections;
using System.Collections.Generic;
using ShapeCrawler.Exceptions;

namespace ShapeCrawler.Charts;

internal class NullCategories : IReadOnlyList<ICategory>
{
    private const string error =
        $"Chart does not have categories. Use {nameof(IChart.HasCategories)} property to check if chart categories are available.";

    public IEnumerator<ICategory> GetEnumerator() => throw new SCException(error);
    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();
    public int Count => throw new SCException(error);
    public ICategory this[int index] => throw new SCException(error);
}
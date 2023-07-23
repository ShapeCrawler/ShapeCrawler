using System;

namespace ShapeCrawler.Shared;

internal sealed class ResetAbleLazy<T>
{
    private readonly Func<T> valueFactory;
    private Lazy<T> lazy;

    internal ResetAbleLazy(Func<T> valueFactory)
    {
        this.valueFactory = valueFactory;
        this.lazy = new Lazy<T>(this.valueFactory);
    }

    internal T Value => this.lazy.Value;

    internal void Reset()
    {
        this.lazy = new Lazy<T>(this.valueFactory);
    }
}
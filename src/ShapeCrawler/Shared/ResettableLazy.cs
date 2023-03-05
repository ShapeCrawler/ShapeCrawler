using System;

namespace ShapeCrawler.Shared;

internal sealed class ResettableLazy<T>
{
    private readonly Func<T?> valueFactory;
    private Lazy<T> lazy;

    public ResettableLazy(Func<T?> valueFactory)
    {
        this.valueFactory = valueFactory;
        this.lazy = new Lazy<T>(this.valueFactory);
    }

    public T Value => this.lazy.Value;

    public void Reset()
    {
        this.lazy = new Lazy<T>(this.valueFactory);
    }
}
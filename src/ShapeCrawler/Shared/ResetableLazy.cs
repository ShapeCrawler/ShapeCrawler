using System;

namespace ShapeCrawler.Shared;

internal sealed class ResetableLazy<T>
{
    private readonly Func<T> valueFactory;
    private Lazy<T> lazy;

    internal ResetableLazy(Func<T> valueFactory)
    {
        this.valueFactory = valueFactory;
        this.lazy = new Lazy<T>(this.valueFactory);
    }

    internal T Value => this.lazy.Value;

    internal void Reset()
    {
        if (this.lazy.IsValueCreated)
        {
            this.lazy = new Lazy<T>(this.valueFactory);
        }
    }
}
using System;
using System.Threading;

namespace ShapeCrawler.Shared
{
    internal class ResettableLazy<T>
    {
        private Lazy<T> _lazy;
        private readonly Func<T> _valueFactory;

        public T Value => _lazy.Value;

        public ResettableLazy(Func<T> valueFactory)
        {
            _valueFactory = valueFactory;
            _lazy = new Lazy<T>(_valueFactory);
        }

        public void Reset()
        {
            _lazy = new Lazy<T>(_valueFactory);
        }
    }
}

using System;

namespace ShapeCrawler.Shared
{
    public class ResettableLazy<T> //TODO: convert in internal
    {
        private readonly Func<T> _valueFactory;
        private Lazy<T> _lazy;

        public ResettableLazy(Func<T> valueFactory)
        {
            _valueFactory = valueFactory;
            _lazy = new Lazy<T>(_valueFactory);
        }

        public T Value => _lazy.Value;

        public void Reset()
        {
            _lazy = new Lazy<T>(_valueFactory);
        }
    }
}
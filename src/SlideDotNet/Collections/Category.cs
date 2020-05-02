using System;

namespace SlideDotNet.Collections
{
    public class Category
    {
        public Category Parent { get; }

        public string Value { get; }

        public Category(string value)
        {
            Value = value ?? throw new ArgumentNullException(nameof(value));
        }

        public Category(string value, Category parent)
        {
            Value = value ?? throw new ArgumentNullException(nameof(value));
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
        }
    }
}
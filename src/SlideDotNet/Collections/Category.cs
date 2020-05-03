using System;

namespace SlideDotNet.Collections
{
    /// <summary>
    /// Represents a chart category.
    /// </summary>
    public class Category
    {
        #region Properties

        /// <summary>
        /// Returns the parent category. Returns null if the chart hash not multi-category.
        /// </summary>
        public Category Parent { get; }

        /// <summary>
        /// Returns category name.
        /// </summary>
        public string Name { get; }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Constructs a non-multi-category.
        /// </summary>
        /// <param name="value"></param>
        public Category(string value)
        {
            Name = value ?? throw new ArgumentNullException(nameof(value));
        }

        /// <summary>
        /// Constructs a multi-category.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="parent"></param>
        public Category(string value, Category parent)
        {
            Name = value ?? throw new ArgumentNullException(nameof(value));
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
        }

        #endregion Constructors
    }
}
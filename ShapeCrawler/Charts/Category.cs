using System;

namespace ShapeCrawler.Charts
{
    /// <summary>
    /// Represents a chart category.
    /// </summary>
    public class Category
    {
        #region Properties

        /// <summary>
        /// Gets main category. Returns null if the chart is not a multi-category chart type.
        /// </summary>
        public Category MainCategory { get; }

        /// <summary>
        /// Gets category name.
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
            MainCategory = parent ?? throw new ArgumentNullException(nameof(parent));
        }

        #endregion Constructors
    }
}
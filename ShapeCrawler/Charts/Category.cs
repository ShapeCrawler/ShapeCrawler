using System;

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents a chart category.
    /// </summary>
    public class Category
    {
        #region Properties

        /// <summary>
        ///     Gets main category. Returns <c>NULL</c> if the chart is not a multi-category chart type.
        /// </summary>
        public Category MainCategory { get; }

        /// <summary>
        ///     Gets category name.
        /// </summary>
        public string Name { get; }

        #endregion Properties

        #region Constructors

        /// <summary>
        ///     Initializes a new non-multi-category.
        /// </summary>
        internal Category(string value)
        {
            Name = value ?? throw new ArgumentNullException(nameof(value));
        }

        /// <summary>
        ///     Initializes a new multi-category.
        /// </summary>
        internal Category(string value, Category parent)
        {
            Name = value ?? throw new ArgumentNullException(nameof(value));
            MainCategory = parent ?? throw new ArgumentNullException(nameof(parent));
        }

        #endregion Constructors
    }
}
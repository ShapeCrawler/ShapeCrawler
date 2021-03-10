using System;

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents a chart category.
    /// </summary>
    public class Category
    {
        #region Constructors

        /// <summary>
        ///     Initializes non-multi-category.
        /// </summary>
        internal Category(string name)
        {
            Name = name;
        }

        /// <summary>
        ///     Initializes multi-category.
        /// </summary>
        internal Category(string name, Category main)
        {
            Name = name;
            MainCategory = main;
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        ///     Gets main category. Returns <c>NULL</c> if the chart is not a multi-category chart type.
        /// </summary>
        public Category MainCategory { get; }

#if DEBUG
        /// <summary>
        ///     Gets or sets category name.
        /// </summary>
        public string Name { get; set; }
#else
        /// <summary>
        ///     Gets category name.
        /// </summary>
        public string Name { get; }
#endif

        #endregion Properties
    }
}
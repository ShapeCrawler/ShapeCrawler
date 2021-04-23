﻿using ShapeCrawler.AutoShapes;
using ShapeCrawler.Collections;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a paragraph.
    /// </summary>
    public interface IParagraph
    {
        /// <summary>
        ///     Gets or sets paragraph text.
        /// </summary>
        string Text { get; set; }

        /// <summary>
        ///     Gets collection of paragraph portions.
        /// </summary>
        IPortionCollection Portions { get; }

        /// <summary>
        ///     Gets paragraph bullet. Returns <c>NULL</c> if bullet does not exist.
        /// </summary>
        Bullet Bullet { get; }
    }
}
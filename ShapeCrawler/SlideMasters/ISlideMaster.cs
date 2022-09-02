using System.Collections.Generic;
using ShapeCrawler.Drawing;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a Slide Master.
    /// </summary>
    public interface ISlideMaster
    {
        /// <summary>
        ///     Gets background image.
        /// </summary>
        SCImage Background { get; }

        /// <summary>
        ///     Gets collection of Slide Layouts.
        /// </summary>
        IReadOnlyList<ISlideLayout> SlideLayouts { get; }
        
        IShapeCollection Shapes { get; }
    }
}
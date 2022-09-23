using System.Collections.Generic;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a slide.
    /// </summary>
    public interface ISlide
    {
        /// <summary>
        ///     Gets or sets slide number.
        /// </summary>
        int Number { get; set; }

        /// <summary>
        ///     Gets background image of the slide. Returns <c>NULL</c> if slide does not have background.
        /// </summary>
        SCImage? Background { get; }

        /// <summary>
        ///     Gets or sets custom data.
        /// </summary>
        string CustomData { get; set; }

        /// <summary>
        ///     Gets a value indicating whether the slide is hidden.
        /// </summary>
        bool Hidden { get; }

        /// <summary>
        ///     Gets referenced Slide Layout.
        /// </summary>
        ISlideLayout SlideLayout { get; }

        /// <summary>
        ///     Gets presentation.
        /// </summary>
        IPresentation Presentation { get; }

        /// <summary>
        ///     Gets instance of <see cref=" DocumentFormat.OpenXml.Packaging.SlidePart"/> class of the underlying Open XML SDK.
        /// </summary>
        SlidePart SDKSlidePart { get; }
        
        /// <summary>
        ///     Gets collection of shapes.
        /// </summary>
        IShapeCollection Shapes { get; }

        /// <summary>
        /// Gets a list of all textboxes on that slide, including those in tables.
        /// </summary>
        public IList<ITextFrame> GetAllTextFrames();

        /// <summary>
        ///     Hides slide.
        /// </summary>
        void Hide();

#if DEBUG
        /// <summary>
        ///     Converts slide to HTML.
        /// </summary>
        Task<string> ToHtml();
#endif
    }
}
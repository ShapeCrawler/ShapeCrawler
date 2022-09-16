using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a user slide.
    /// </summary>
    public interface ISlide
    {
        /// <summary>
        ///     Gets or sets slide number.
        /// </summary>
        int Number { get; set; }

        /// <summary>
        ///     Gets background image of the slide. Returns <c>NULL</c> if the slide does not have background.
        /// </summary>
        SCImage Background { get; }

        /// <summary>
        ///     Gets or sets custom data.
        /// </summary>
        string CustomData { get; set; }

        /// <summary>
        ///     Gets a value indicating whether slide hidden.
        /// </summary>
        bool Hidden { get; }

        /// <summary>
        ///     Gets parent (referenced) Slide Layout.
        /// </summary>
        ISlideLayout SlideLayout { get; }

        IPresentation ParentPresentation { get; }

        SlidePart SDKSlidePart { get; }
        
        IShapeCollection Shapes { get; }

        /// <summary>
        /// Gets a list of all textboxes on that slide, including those in tables
        /// </summary>
        public IList<ITextBox> Textboxes { get; }

        /// <summary>
        ///     Hides slide.
        /// </summary>
        void Hide();

        /// <summary>
        ///     Saves slide scheme to stream.
        /// </summary>
        void SaveScheme(Stream stream);

        /// <summary>
        ///     Saves slide scheme to file.
        /// </summary>
        void SaveScheme(string filePath);

#if DEBUG
        /// <summary>
        ///     Converts slide to HTML.
        /// </summary>
        Task<string> ToHtml();
#endif
    }
}
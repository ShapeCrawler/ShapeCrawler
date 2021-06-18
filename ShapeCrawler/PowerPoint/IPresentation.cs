using System;
using System.IO;
using ShapeCrawler.Collections;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a presentation document.
    /// </summary>
    public interface IPresentation : IDisposable
    {
        /// <summary>
        ///     Gets the presentation slides.
        /// </summary>
        ISlideCollection Slides { get; }

        /// <summary>
        ///     Gets the presentation slides width.
        /// </summary>
        int SlideWidth { get; }

        /// <summary>
        ///     Gets the presentation slides height.
        /// </summary>
        int SlideHeight { get; }

        /// <summary>
        ///     Gets collection of the slide masters.
        /// </summary>
        ISlideMasterCollection SlideMasters { get; }

        /// <summary>
        ///     Gets presentation byte array.
        /// </summary>
        byte[] ByteArray { get; }

        /// <summary>
        ///     Saves presentation.
        /// </summary>
        void Save();

        /// <summary>
        ///     Saves presentation in specified file path.
        /// </summary>
        void SaveAs(string filePath);

        /// <summary>
        ///     Saves presentation in specified stream.
        /// </summary>
        void SaveAs(Stream stream);

        /// <summary>
        ///     Saves and closes the presentation.
        /// </summary>
        void Close();

        void Dispose();
    }
}
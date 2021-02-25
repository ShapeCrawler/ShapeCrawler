using System;
using System.IO;
using ShapeCrawler.Collections;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler
{
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

        SlideMasterCollection SlideMasters { get; }
        void Save();

        /// <summary>
        ///     Saves presentation in specified file path.
        /// </summary>
        /// <param name="filePath"></param>
        void SaveAs(string filePath);

        /// <summary>
        ///     Saves presentation in specified stream.
        /// </summary>
        /// <param name="stream"></param>
        void SaveAs(Stream stream);

        /// <summary>
        ///     Saves and closes the presentation.
        /// </summary>
        void Close();

        void Dispose();
    }
}
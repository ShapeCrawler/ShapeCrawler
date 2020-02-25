using System;

namespace SlideDotNet.Models
{
    /// <summary>
    /// Represents a presentation.
    /// </summary>
    public interface IPresentation : IDisposable
    {
        /// <summary>
        /// Returns slides collection.
        /// </summary>
        ISlideCollection Slides { get; }

        /// <summary>
        /// Returns slides width in EMUs.
        /// </summary>
        int SlideWidth { get; }

        /// <summary>
        /// Returns slides height in EMUs.
        /// </summary>
        int SlideHeight { get; }

        /// <summary>
        /// Saves the presentation in specified file path.
        /// </summary>
        /// <param name="filePath"></param>
        void SaveAs(string filePath);

        /// <summary>
        /// Saves and closes the current presentation.
        /// </summary>
        void Close();
    }
}
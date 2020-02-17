using System;

namespace SlideDotNet.Models
{
    /// <summary>
    /// Provides APIs for presentation document.
    /// </summary>
    public interface IPresentation : IDisposable
    {
        ISlideCollection Slides { get; }

        void SaveAs(string filePath);
    }
}
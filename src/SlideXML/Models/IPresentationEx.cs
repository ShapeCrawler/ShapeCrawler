using System;

namespace SlideXML.Models
{
    /// <summary>
    /// Provides APIs for presentation document.
    /// </summary>
    public interface IPresentationEx : IDisposable
    {
        ISlideCollection Slides { get; }

        void SaveAs(string filePath);
    }
}
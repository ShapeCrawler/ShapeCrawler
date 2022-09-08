// ReSharper disable CheckNamespace

using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents base class for Slide, Slide Layout and Slide Master.
    /// </summary>
    internal abstract class SlideBase : IRemovable
    {
        public abstract bool IsRemoved { get; set; }

        public abstract void ThrowIfRemoved();
        
        internal abstract OpenXmlPart OpenXmlPart { get; }
    }
}
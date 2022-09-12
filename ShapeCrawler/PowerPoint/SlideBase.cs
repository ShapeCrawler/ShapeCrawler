using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Services;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents base class for Slide, Slide Layout and Slide Master.
    /// </summary>
    internal abstract class SlideBase : IRemovable
    {
        public abstract bool IsRemoved { get; set; }

        internal abstract TypedOpenXmlPart TypedOpenXmlPart { get; }

        public abstract void ThrowIfRemoved();
    }
}
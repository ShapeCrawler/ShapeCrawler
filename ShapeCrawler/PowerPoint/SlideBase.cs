using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Services;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents base class for Slide, Slide Layout and Slide Master.
    /// </summary>
    internal abstract class SlideBase : IRemovable, IPresentationComponent
    {
        public abstract bool IsRemoved { get; set; } // TODO: make internal
        public SCPresentation PresentationInternal { get; set; } // TODO: make internal

        internal abstract TypedOpenXmlPart TypedOpenXmlPart { get; }
        
        

        public abstract void ThrowIfRemoved(); // TODO: make internal
    }
}
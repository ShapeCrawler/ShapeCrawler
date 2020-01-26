using DocumentFormat.OpenXml.Packaging;
using SlideXML.Models;

namespace SlideXML.Services
{
    /// <summary>
    /// Provides APIs to parse background images.
    /// </summary>
    public interface IBackgroundImageFactory
    {
        ImageEx CreateBackgroundSlide(SlidePart sldPart);

        ImageEx CreateBackgroundShape(SlidePart sldPart, DocumentFormat.OpenXml.Presentation.Shape pShape);
    }
}
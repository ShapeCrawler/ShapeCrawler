using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models;

namespace SlideDotNet.Services
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
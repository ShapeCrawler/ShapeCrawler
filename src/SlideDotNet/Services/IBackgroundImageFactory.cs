using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services
{
    /// <summary>
    /// Represents a factory to create an instance of the <see cref="ImageEx"/> class.
    /// </summary>
    public interface IBackgroundImageFactory
    {
        /// <summary>
        /// Gets slide background image. Returns null if slide does not have background image.
        /// </summary>
        ImageEx FromXmlSlide(SlidePart xmlSldPart);

        /// <summary>
        /// Gets shape background image. It returns null if the shape is not filled with a picture.
        /// </summary>
        ImageEx FromXmlShape(SlidePart xmlSldPart, P.Shape xmlShape);
    }
}
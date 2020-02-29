using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models;

namespace SlideDotNet.Services
{
    /// <summary>
    /// Represents a factory to create an instance of the <see cref="ImageEx"/> class.
    /// </summary>
    public interface IImageExFactory
    {
        /// <summary>
        /// Gets slide background image. Returns <c>null</c> if slide does not have background image.
        /// </summary>
        ImageEx TryFromXmlSlide(SlidePart xmlSldPart);

        /// <summary>
        /// Gets shape background image. Returns <c>null</c> if the shape is not filled with a picture.
        /// </summary>
        ImageEx TryFromXmlShape(SlidePart xmlSldPart, OpenXmlCompositeElement ce);
    }
}
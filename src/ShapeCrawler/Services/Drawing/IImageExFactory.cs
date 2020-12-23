using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Services.Drawing
{
    /// <summary>
    /// Represents a factory to create an instance of the <see cref="ImageEx"/> class.
    /// </summary>
    public interface IImageExFactory
    {
        /// <summary>
        /// Gets slide background image. Returns <c>null</c> if slide does not have background image.
        /// </summary>
        ImageEx TryFromSdkSlide(SlidePart xmlSldPart);

        /// <summary>
        /// Gets shape background image. Returns <c>null</c> if the shape is not filled with a picture.
        /// </summary>
        ImageEx TryFromSdkShape(SlidePart xmlSldPart, OpenXmlCompositeElement ce);
    }
}
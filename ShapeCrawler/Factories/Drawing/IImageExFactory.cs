using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Factories.Drawing
{
    /// <summary>
    /// Represents a factory to create an instance of the <see cref="ImageSc"/> class.
    /// </summary>
    public interface IImageExFactory
    {
        /// <summary>
        /// Gets slide background image. Returns <c>null</c> if slide does not have background image.
        /// </summary>
        ImageSc TryFromSdkSlide(SlidePart xmlSldPart);

        /// <summary>
        /// Gets shape background image. Returns <c>null</c> if the shape is not filled with a picture.
        /// </summary>
        ImageSc TryFromSdkShape(SlidePart xmlSldPart, OpenXmlCompositeElement ce);
    }
}
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using Shape = ShapeCrawler.Models.SlideComponents.Shape;
using Slide = ShapeCrawler.Models.Slide;

namespace ShapeCrawler.Factories.ShapeCreators
{
    /// <summary>
    /// Represents a factory to generate instances of the <see cref="Shape"/> class.
    /// </summary>
    /// <remarks>
    /// <see cref="P.ShapeTree"/> and <see cref="P.GroupShape"/> both derived from <see cref="P.GroupShapeType"/> class.
    /// </remarks>
    public interface IShapeFactory
    {
        /// <summary>
        /// Creates collection of the shapes from SDK-slide part.
        /// </summary>
        /// <param name="sdkSldPart"></param>
        /// <param name="slide"></param>
        /// <returns></returns>
        IList<Shape> FromSdlSlidePart(SlidePart sdkSldPart, Slide slide);
    }
}
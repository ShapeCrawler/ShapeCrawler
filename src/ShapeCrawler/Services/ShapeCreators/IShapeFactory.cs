using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Services.ShapeCreators
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
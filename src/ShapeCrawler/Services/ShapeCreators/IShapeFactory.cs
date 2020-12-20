using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models.SlideComponents;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services.ShapeCreators
{
    /// <summary>
    /// Represents a factory to generate instances of the <see cref="ShapeEx"/> class.
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
        /// <returns></returns>
        IList<ShapeEx> FromSldPart(SlidePart sdkSldPart);
    }
}
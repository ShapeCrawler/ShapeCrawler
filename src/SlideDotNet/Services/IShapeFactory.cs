using System.Collections.Generic;
using SlideDotNet.Models.SlideComponents;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services
{
    /// <summary>
    /// Represents a factory to generate instances of the <see cref="ShapeEx"/> class.
    /// </summary>
    /// <remarks>
    /// <see cref="P.ShapeTree"/> and <see cref="P.GroupShape"/> both derived from <see cref="P.GroupShapeType"/> class.
    /// </remarks>
    public interface IShapeFactory
    {
        IList<ShapeEx> FromTree(P.ShapeTree sdkShapeTree);
    }
}
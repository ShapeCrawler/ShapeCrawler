using ShapeCrawler.Enums;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Services
{
    /// <summary>
    /// Represents a factory to generate a shape geometry.
    /// </summary>
    public interface IGeometryFactory
    {
        GeometryType ForShape(P.Shape sdkShape);

        GeometryType ForPicture(P.Picture sdkPicture);
    }
}
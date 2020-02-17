using DocumentFormat.OpenXml;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services.Builders
{
    /// <summary>
    /// Represents a shape builder.
    /// </summary>
    public interface IShapeBuilder
    {
        Shape BuildAutoShape(OpenXmlCompositeElement xmlElement, ElementSettings elSettings);

        Shape BuildPicture(OpenXmlCompositeElement xmlElement, ElementSettings elSettings);

        Shape BuildTable(P.GraphicFrame xmlGrFrame, ElementSettings elSettings);

        Shape BuildChart(P.GraphicFrame xmlGrFrame);

        Shape BuildOleObject(OpenXmlCompositeElement xmlElement);

        Shape BuildGroup(IElementFactory elFactory, OpenXmlCompositeElement xmlElement, IParents parents);
    }
}

using DocumentFormat.OpenXml;
using SlideXML.Models.Settings;
using SlideXML.Models.SlideComponents;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Services.Builders
{
    /// <summary>
    /// Defines method to create <see cref="ShapeSL"/> instance.
    /// </summary>
    public interface IShapeBuilder
    {
        ShapeSL BuildAutoShape(OpenXmlCompositeElement compositeElement, ElementSettings spSettings);

        ShapeSL BuildChart(P.GraphicFrame xmlGrFrame);

        ShapeSL BuildTable(P.GraphicFrame xmlGrFrame, ElementSettings elSettings);

        ShapeSL BuildPicture(OpenXmlCompositeElement ce);

        ShapeSL BuildGroup(IElementFactory elFactory, OpenXmlCompositeElement compositeElement, IPreSettings preSettings);
        
        ShapeSL BuildOLEObject(OpenXmlCompositeElement compositeElement);
    }
}

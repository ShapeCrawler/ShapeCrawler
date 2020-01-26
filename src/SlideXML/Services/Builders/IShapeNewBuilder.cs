using DocumentFormat.OpenXml;
using SlideXML.Models.Elements;
using SlideXML.Models.Settings;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Services.Builders
{
    /// <summary>
    /// Defines method to create <see cref="ShapeSL"/> instance.
    /// </summary>
    public interface IShapeNewBuilder
    {
        ShapeSL BuildTxtShape(OpenXmlCompositeElement compositeElement, ElementSettings spSettings);

        ShapeSL BuildChartShape(P.GraphicFrame xmlGrFrame);

        ShapeSL BuildTableShape(P.GraphicFrame xmlGrFrame, ElementSettings elSettings);

        ShapeSL BuildPictureShape(OpenXmlCompositeElement ce);

        ShapeSL BuildGroupShape(IElementFactory elFactory, OpenXmlCompositeElement compositeElement, IPreSettings preSettings);
        
        ShapeSL BuildOLEObject(OpenXmlCompositeElement compositeElement);
    }
}

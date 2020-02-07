using DocumentFormat.OpenXml;
using SlideXML.Models.Settings;
using SlideXML.Models.SlideComponents;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Services.Builders
{
    /// <summary>
    /// Defines method to create <see cref="SlideElement"/> instance.
    /// </summary>
    public interface IShapeBuilder
    {
        SlideElement BuildAutoShape(OpenXmlCompositeElement compositeElement, ElementSettings spSettings);

        SlideElement BuildChart(P.GraphicFrame xmlGrFrame);

        SlideElement BuildTable(P.GraphicFrame xmlGrFrame, ElementSettings elSettings);

        SlideElement BuildPicture(OpenXmlCompositeElement ce);

        SlideElement BuildGroup(IElementFactory elFactory, OpenXmlCompositeElement compositeElement, IPreSettings preSettings);
        
        SlideElement BuildOLEObject(OpenXmlCompositeElement compositeElement);
    }
}

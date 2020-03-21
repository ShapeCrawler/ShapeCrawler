using DocumentFormat.OpenXml;
using SlideDotNet.Models.SlideComponents;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services
{
    public abstract class OpenXmlElementHandler
    {
        public OpenXmlElementHandler Successor { get; set; }
        
        public abstract ShapeEx Create(OpenXmlElement openXmlElement);
    }
}

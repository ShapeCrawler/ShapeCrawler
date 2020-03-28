using DocumentFormat.OpenXml;
using SlideDotNet.Models.SlideComponents;

namespace SlideDotNet.Services.ShapeCreators
{
    public abstract class OpenXmlElementHandler
    {
        public OpenXmlElementHandler Successor { get; set; }
        
        public abstract ShapeEx Create(OpenXmlElement openXmlElement);
    }
}

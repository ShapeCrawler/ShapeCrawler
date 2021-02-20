using DocumentFormat.OpenXml;
using ShapeCrawler.Factories.Placeholders;

namespace ShapeCrawler.Placeholders
{
    public class Placeholder
    {
        private readonly OpenXmlCompositeElement _pShapeTreeChild;

        internal Placeholder(OpenXmlCompositeElement pShapeTreeChild)
        {
            _pShapeTreeChild = pShapeTreeChild;
        }

        public PlaceholderType Type => PlaceholderService.GetPlaceholderType(_pShapeTreeChild);
    }
}
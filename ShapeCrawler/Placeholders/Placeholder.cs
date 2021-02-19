using DocumentFormat.OpenXml;

namespace ShapeCrawler.Factories.Placeholders
{
    public class Placeholder
    {
        private readonly OpenXmlCompositeElement _shapeTreeSource;

        internal Placeholder(OpenXmlCompositeElement shapeTreeSource)
        {
            _shapeTreeSource = shapeTreeSource;
        }

        public PlaceholderType Type => PlaceholderService.GetPlaceholderType(_shapeTreeSource);
    }
}
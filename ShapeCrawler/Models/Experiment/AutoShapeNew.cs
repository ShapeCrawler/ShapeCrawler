using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Factories.Placeholders;

namespace ShapeCrawler.Models.Experiment
{
    public class AutoShapeNew
    {
        private readonly Shape _pShape;

        public uint Id => _pShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id;
        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; }
        public GeometryType GeometryType { get; }
        public Placeholder Placeholder { get; }
    }
}
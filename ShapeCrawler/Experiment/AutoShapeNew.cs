using ShapeCrawler.Placeholders;

namespace ShapeCrawler.Experiment
{
    internal class AutoShapeNew
    {
        private readonly DocumentFormat.OpenXml.Presentation.Shape _pShape;

        public uint Id => _pShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id;
        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; }
        public GeometryType GeometryType { get; }
        public Placeholder Placeholder { get; }
    }
}
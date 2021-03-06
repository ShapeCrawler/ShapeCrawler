using System;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler.Collections
{
    internal class LayoutChart : IShape
    {
        public LayoutChart(SlideLayoutSc slideLayout, DocumentFormat.OpenXml.Presentation.GraphicFrame pGraphicFrame)
        {
            throw new NotImplementedException();
        }

        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }
        public int Id { get; }
        public string Name { get; }
        public bool Hidden { get; }
        public IPlaceholder Placeholder { get; }
        public GeometryType GeometryType { get; }
        public string CustomData { get; set; }
    }
}
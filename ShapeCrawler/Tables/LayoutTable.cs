using System;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections
{
    internal class LayoutTable : LayoutShape, IShape
    {
        public LayoutTable(SlideLayoutSc slideLayout, P.GraphicFrame pGraphicFrame) : base(slideLayout, pGraphicFrame)
        {
        }

        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }
        public int Id { get; }
        public string Name { get; }
        public bool Hidden { get; }
        public GeometryType GeometryType { get; }
    }
}
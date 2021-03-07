using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tables
{
    internal class LayoutTable : LayoutShape, IShape
    {
        public LayoutTable(SlideLayoutSc slideLayout, P.GraphicFrame pGraphicFrame) : base(slideLayout, pGraphicFrame)
        {
        }

        public string Name { get; }
        public bool Hidden { get; }
    }
}
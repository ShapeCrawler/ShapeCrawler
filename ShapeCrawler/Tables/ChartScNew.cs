using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Models;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler.Tables
{
    public class ChartScNew : BaseShape
    {
        private readonly SlideMasterSc slideMaster;
        private readonly GraphicFrame pGraphicFrame;

        public ChartScNew(SlideMasterSc slideMaster, GraphicFrame pGraphicFrame)
        {
            this.slideMaster = slideMaster;
            this.pGraphicFrame = pGraphicFrame;
        }

        public override long X { get; }
        public override long Y { get; }
        public override long Width { get; }
        public override long Height { get; }
        public override GeometryType GeometryType { get; }
    }
}
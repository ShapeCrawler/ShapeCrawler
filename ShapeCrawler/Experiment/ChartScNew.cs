using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler.Experiment
{
    public class ChartScNew : BaseShape
    {
        private readonly GraphicFrame pGraphicFrame;
        private readonly SlideMasterSc slideMaster;

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
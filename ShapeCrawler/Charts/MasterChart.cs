using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents a chart on a Slide Master.
    /// </summary>
    internal class MasterChart : MasterShape, IShape
    {
        internal MasterChart(SCSlideMaster slideMaster, P.GraphicFrame pGraphicFrame)
            : base(slideMaster, pGraphicFrame)
        {
        }
    }
}
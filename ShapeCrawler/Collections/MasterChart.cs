using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a chart on a Slide Master.
    /// </summary>
    internal class MasterChart : MasterShape, IShape
    {
        internal MasterChart(SlideMasterSc slideMaster, P.GraphicFrame pGraphicFrame)
            : base(slideMaster, pGraphicFrame)
        {
        }
    }
}
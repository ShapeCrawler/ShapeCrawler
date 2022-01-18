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
        internal MasterChart(SCSlideMaster slideLayoutInternalMaster, P.GraphicFrame pGraphicFrame)
            : base(pGraphicFrame, slideLayoutInternalMaster)
        {
        }

        public override SCSlideMaster ParentSlideMaster { get; set; }
    }
}
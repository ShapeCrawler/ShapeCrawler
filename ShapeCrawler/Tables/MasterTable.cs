using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tables
{
    /// <summary>
    ///     Represents a table on a Slide Master.
    /// </summary>
    internal class MasterTable : MasterShape, IShape
    {
        internal MasterTable(SCSlideMaster slideMasterInternal, P.GraphicFrame pGraphicFrame)
            : base(pGraphicFrame, slideMasterInternal)
        {
        }

        public SCShapeType ShapeType => SCShapeType.Table;

        public override SCPresentation PresentationInternal { get; }
    }
}
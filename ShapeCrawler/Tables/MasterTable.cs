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
        internal MasterTable(P.GraphicFrame pGraphicFrame, SCSlideMaster parentSlideMaster)
            : base(pGraphicFrame, parentSlideMaster)
        {
        }
    }
}
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tables
{
    /// <summary>
    ///     Represents a table on a Slide Master.
    /// </summary>
    internal class MasterTable : MasterShape, IShape
    {
        internal MasterTable(SlideMasterSc slideMaster, P.GraphicFrame pGraphicFrame)
            : base(slideMaster, pGraphicFrame)
        {
        }
    }
}
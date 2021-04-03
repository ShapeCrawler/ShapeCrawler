using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tables
{
    /// <summary>
    ///     Represents a table on a Slide Master.
    /// </summary>
    internal class MasterTable : MasterShape, IShape
    {
        internal MasterTable(SCSlideMaster slideMaster, P.GraphicFrame pGraphicFrame)
            : base(slideMaster, pGraphicFrame)
        {
        }
    }
}
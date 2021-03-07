using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections
{
    internal class MasterTable : MasterShape, IShape
    {
        public MasterTable(SlideMasterSc slideMaster, P.GraphicFrame pGraphicFrame) : base(slideMaster, pGraphicFrame)
        {
        }

        public string Name { get; }
        public bool Hidden { get; }
    }
}
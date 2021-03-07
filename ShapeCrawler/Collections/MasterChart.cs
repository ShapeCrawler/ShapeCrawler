using System;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections
{
    internal class MasterChart : MasterShape, IShape
    {
        public MasterChart(SlideMasterSc slideMaster, P.GraphicFrame pGraphicFrame)
            : base(slideMaster, pGraphicFrame)
        {
        }
        public string Name { get; }
        public bool Hidden { get; }
    }
}
using System;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections
{
    internal class MasterOLEObject : MasterShape, IShape
    {
        private P.GraphicFrame pGraphicFrame;
        private SlideMasterSc slideMaster;

        public MasterOLEObject(SlideMasterSc slideMaster, P.GraphicFrame pGraphicFrame) : base(slideMaster, pGraphicFrame)
        {
            this.slideMaster = slideMaster;
            this.pGraphicFrame = pGraphicFrame;
        }

        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }
        public int Id { get; }
        public string Name { get; }
        public bool Hidden { get; }
        public override IPlaceholder Placeholder => throw new NotImplementedException();
        public GeometryType GeometryType { get; }
    }
}
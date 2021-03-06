using System;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections
{
    internal class MasterOLEObject : MasterShape
    {
        private SlideMasterSc slideMaster;
        private P.GraphicFrame pGraphicFrame;

        public MasterOLEObject(SlideMasterSc slideMaster, P.GraphicFrame pGraphicFrame) : base(slideMaster, pGraphicFrame)
        {
            this.slideMaster = slideMaster;
            this.pGraphicFrame = pGraphicFrame;
        }

        public override IPlaceholder Placeholder => throw new NotImplementedException();
    }
}
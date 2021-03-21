﻿using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.OLEObjects
{
    internal class MasterOLEObject : MasterShape, IShape
    {
        public MasterOLEObject(SlideMasterSc slideMaster, P.GraphicFrame pGraphicFrame) : base(slideMaster,
            pGraphicFrame)
        {
        }
        public override IPlaceholder Placeholder => MasterPlaceholder.Create(PShapeTreeChild);
    }
}
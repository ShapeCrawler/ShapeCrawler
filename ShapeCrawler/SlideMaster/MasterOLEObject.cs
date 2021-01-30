using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.SlideMaster;

// ReSharper disable CheckNamespace

namespace ShapeCrawler
{
    public class MasterOLEObject : MasterShape
    {
        public MasterOLEObject(SlideMasterSc slideMasterSc, GraphicFrame pGraphicFrame) 
            : base(slideMasterSc, pGraphicFrame)
        {

        }
    }
}
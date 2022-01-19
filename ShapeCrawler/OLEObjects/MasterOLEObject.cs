using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.OLEObjects
{
    internal class MasterOLEObject : MasterShape, IShape
    {
        public MasterOLEObject(P.GraphicFrame pGraphicFrame, SCSlideMaster parentSlideLayoutInternalMaster) 
            : base(pGraphicFrame, parentSlideLayoutInternalMaster)
        {
        }

        public override IPlaceholder Placeholder => MasterPlaceholder.Create(PShapeTreesChild);
        public override SCSlideMaster ParentSlideMaster { get; set; }
    }
}
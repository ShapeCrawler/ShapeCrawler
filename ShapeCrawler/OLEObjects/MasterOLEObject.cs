using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.OLEObjects
{
    internal class MasterOLEObject : MasterShape, IShape
    {
        public MasterOLEObject(P.GraphicFrame pGraphicFrame, SCSlideMaster slideMasterInternal)
            : base(pGraphicFrame, slideMasterInternal)
        {
        }

        public override IPlaceholder Placeholder => MasterPlaceholder.Create(PShapeTreesChild);

        public ShapeType ShapeType => ShapeType.OLEObject;
    }
}
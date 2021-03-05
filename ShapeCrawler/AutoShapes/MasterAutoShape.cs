using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Settings;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler
{
    internal class MasterAutoShape : AutoShape
    {
        public override IPlaceholder Placeholder => MasterPlaceholder.Create(this);

        internal MasterAutoShape(ILocation innerTransform, ShapeContext spContext, GeometryType geometryType, DocumentFormat.OpenXml.Presentation.Shape pShape, SlideSc slide) 
            : base(innerTransform, spContext, geometryType, pShape, slide)
        {
        }

        internal MasterAutoShape(SlideLayoutSc slideLayout, DocumentFormat.OpenXml.Presentation.Shape pShape) 
            : base(slideLayout, pShape)
        {
        }
    }
}
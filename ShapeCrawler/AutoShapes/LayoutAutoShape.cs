using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Settings;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler
{
    internal class LayoutAutoShape : AutoShape
    {
        public override IPlaceholder Placeholder => LayoutPlaceholder.Create(this);

        internal LayoutAutoShape(ILocation innerTransform, ShapeContext spContext, GeometryType geometryType, DocumentFormat.OpenXml.Presentation.Shape pShape, SlideSc slide) 
            : base(innerTransform, spContext, geometryType, pShape, slide)
        {
        }

        internal LayoutAutoShape(SlideLayoutSc slideLayout, DocumentFormat.OpenXml.Presentation.Shape pShape) : base(slideLayout, pShape)
        {
        }
    }
}
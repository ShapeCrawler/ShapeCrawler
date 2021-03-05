using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Settings;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler
{
    internal class SlideAutoShape : AutoShape
    {
        public override IPlaceholder Placeholder => SlidePlaceholder.Create(this);

        internal SlideAutoShape(ILocation innerTransform, ShapeContext spContext, GeometryType geometryType, DocumentFormat.OpenXml.Presentation.Shape pShape, SlideSc slide) 
            : base(innerTransform, spContext, geometryType, pShape, slide)
        {
        }

        internal SlideAutoShape(SlideLayoutSc slideLayout, DocumentFormat.OpenXml.Presentation.Shape pShape) : base(slideLayout, pShape)
        {
        }
    }
}
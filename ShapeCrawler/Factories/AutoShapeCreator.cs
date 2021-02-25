using DocumentFormat.OpenXml;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class AutoShapeCreator : OpenXmlElementHandler
    {
        #region Constructors

        public AutoShapeCreator(ShapeContext.Builder shapeContextBuilder,
            LocationParser transformFactory,
            GeometryFactory geometryFactory)
        {
            _shapeContextBuilder = shapeContextBuilder;
            _transformFactory = transformFactory;
            _geometryFactory = geometryFactory;
        }

        #endregion Constructors

        #region Public Methods

        public override IShape Create(OpenXmlCompositeElement pShapeTreeChild, SlideSc slide)
        {
            if (pShapeTreeChild is P.Shape pShape)
            {
                ShapeContext shapeContext = _shapeContextBuilder.Build(pShapeTreeChild);
                ILocation innerTransform = _transformFactory.FromComposite(pShape);
                GeometryType geometryType = _geometryFactory.ForCompositeElement(pShape, pShape.ShapeProperties);
                var autoShape = new AutoShape(innerTransform, shapeContext, geometryType, pShape, slide);

                return autoShape;
            }

            return Successor?.Create(pShapeTreeChild, slide);
        }

        #endregion Public Methods

        #region Fields

        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly LocationParser _transformFactory;
        private readonly GeometryFactory _geometryFactory;

        #endregion Fields
    }
}
using DocumentFormat.OpenXml;
using ShapeCrawler.Factories;
using ShapeCrawler.Factories.ShapeCreators;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Settings;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes
{
    /// <summary>
    /// <inheritdoc cref="OpenXmlElementHandler"/>.
    /// </summary>
    internal class AutoShapeCreator : OpenXmlElementHandler
    {
        #region Fields

        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly LocationParser _transformFactory;
        private readonly GeometryFactory _geometryFactory;

        #endregion Fields

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

        public override IShape Create(OpenXmlCompositeElement shapeTreeSource, SlideSc slide)
        {
            if (shapeTreeSource is P.Shape pShape)
            {
                ShapeContext shapeContext = _shapeContextBuilder.Build(shapeTreeSource);
                ILocation innerTransform = _transformFactory.FromComposite(pShape);
                GeometryType geometryType = _geometryFactory.ForCompositeElement(pShape, pShape.ShapeProperties);
                var autoShape = new AutoShape(innerTransform, shapeContext, geometryType, shapeTreeSource, slide);
                
                return autoShape;
            }

            return Successor?.Create(shapeTreeSource, slide);
        }

        #endregion Public Methods
    }
}
using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Factories.Builders;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories.ShapeCreators
{
    /// <summary>
    /// <inheritdoc cref="OpenXmlElementHandler"/>.
    /// </summary>
    internal class PShapeHandler : OpenXmlElementHandler
    {
        #region Fields

        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly LocationParser _transformFactory;
        private readonly GeometryFactory _geometryFactory;
        private readonly IShapeBuilder _shapeBuilder;

        #endregion Fields

        #region Constructors

        public PShapeHandler(ShapeContext.Builder shapeContextBuilder,
                               LocationParser transformFactory,
                               GeometryFactory geometryFactory) :
            this(shapeContextBuilder, transformFactory, geometryFactory, new ShapeSc.Builder())
        {

        }

        public PShapeHandler(ShapeContext.Builder shapeContextBuilder,
                               LocationParser transformFactory,
                               GeometryFactory geometryFactory,
                               IShapeBuilder shapeBuilder)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
            _geometryFactory = geometryFactory ?? throw new ArgumentNullException(nameof(geometryFactory));
            _shapeBuilder = shapeBuilder ?? throw new ArgumentNullException(nameof(shapeBuilder));
        }

        #endregion Constructors

        #region Public Methods

        public override ShapeSc Create(OpenXmlCompositeElement shapeTreeSource)
        {
            Check.NotNull(shapeTreeSource, nameof(shapeTreeSource));

            if (shapeTreeSource is P.Shape pShape)
            {
                ShapeContext shapeContext = _shapeContextBuilder.Build(shapeTreeSource);
                var innerTransform = _transformFactory.FromComposite(pShape);
                var geometry = _geometryFactory.ForCompositeElement(pShape, pShape.ShapeProperties);
                var shape = _shapeBuilder.WithAutoShape(innerTransform, shapeContext, geometry, shapeTreeSource);
                
                return shape;
            }
            
            if (Successor != null)
            {
                return Successor.Create(shapeTreeSource);
            }
           
            return null;
        }

        #endregion Public Methods
    }
}
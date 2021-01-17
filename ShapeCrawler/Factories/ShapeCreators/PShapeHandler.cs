using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Factories.Builders;
using ShapeCrawler.Models.SlideComponents;
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
        private readonly IGeometryFactory _geometryFactory;
        private readonly IShapeBuilder _shapeBuilder;

        #endregion Fields

        #region Constructors

        public PShapeHandler(ShapeContext.Builder shapeContextBuilder,
                               LocationParser transformFactory,
                               IGeometryFactory geometryFactory) :
            this(shapeContextBuilder, transformFactory, geometryFactory, new ShapeSc.Builder())
        {

        }

        //TODO: inject interface instead
        public PShapeHandler(ShapeContext.Builder shapeContextBuilder,
                               LocationParser transformFactory,
                               IGeometryFactory geometryFactory,
                               IShapeBuilder shapeBuilder)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
            _geometryFactory = geometryFactory ?? throw new ArgumentNullException(nameof(geometryFactory));
            _shapeBuilder = shapeBuilder ?? throw new ArgumentNullException(nameof(shapeBuilder));
        }

        #endregion Constructors

        #region Public Methods

        public override ShapeSc Create(OpenXmlElement sdkElement)
        {
            Check.NotNull(sdkElement, nameof(sdkElement));

            if (sdkElement is P.Shape pShape)
            {
                var spContext = _shapeContextBuilder.Build(sdkElement);
                var innerTransform = _transformFactory.FromComposite(pShape);
                var geometry = _geometryFactory.ForShape(pShape);
                var shape = _shapeBuilder.WithAutoShape(innerTransform, spContext, geometry);
                
                return shape;
            }
            
            if (Successor != null)
            {
                return Successor.Create(sdkElement);
            }
           
            return null;
        }

        #endregion Public Methods
    }
}
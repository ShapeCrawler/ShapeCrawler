using System;
using DocumentFormat.OpenXml;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Services.Builders;
using SlideDotNet.Shared;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Services.ShapeCreators
{
    /// <summary>
    /// <inheritdoc cref="OpenXmlElementHandler"/>.
    /// </summary>
    public class SdkShapeHandler : OpenXmlElementHandler
    {
        #region Fields

        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly LocationParser _transformFactory;
        private readonly IGeometryFactory _geometryFactory;
        private readonly IShapeBuilder _shapeBuilder;

        #endregion Fields

        #region Constructors

        public SdkShapeHandler(ShapeContext.Builder shapeContextBuilder,
                               LocationParser transformFactory,
                               IGeometryFactory geometryFactory) :
            this(shapeContextBuilder, transformFactory, geometryFactory, new ShapeEx.Builder())
        {

        }

        //TODO: inject interface instead
        public SdkShapeHandler(ShapeContext.Builder shapeContextBuilder,
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

        public override ShapeEx Create(OpenXmlElement sdkElement)
        {
            Check.NotNull(sdkElement, nameof(sdkElement));

            if (sdkElement is P.Shape sdkShape)
            {
                var spContext = _shapeContextBuilder.Build(sdkElement);
                var innerTransform = _transformFactory.FromComposite(sdkShape);
                var geometry = _geometryFactory.ForShape(sdkShape);
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
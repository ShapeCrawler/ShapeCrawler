using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Factories.Builders;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories.ShapeCreators
{
    /// <summary>
    /// Represents a picture handler for p:pic and picture p:graphicFrame element.
    /// </summary>
    public class PictureHandler : OpenXmlElementHandler
    {
        #region Fields

        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly LocationParser _transformFactory;
        private readonly IShapeBuilder _shapeBuilder;
        private readonly SlidePart _sdkSldPart;
        private readonly GeometryFactory _geometryFactory;

        #endregion Fields

        #region Constructors

        internal PictureHandler(ShapeContext.Builder shapeContextBuilder,
                              LocationParser transformFactory,
                              GeometryFactory geometryFactory,
                              SlidePart sdkSldPart) :
            this(shapeContextBuilder, transformFactory, geometryFactory, sdkSldPart, new ShapeSc.Builder())
        {

        }

        internal PictureHandler(ShapeContext.Builder shapeContextBuilder,
                              LocationParser transformFactory,
                              GeometryFactory geometryFactory,
                              SlidePart sdkSldPart,
                              IShapeBuilder shapeBuilder)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
            _shapeBuilder = shapeBuilder ?? throw new ArgumentNullException(nameof(shapeBuilder));
            _sdkSldPart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _geometryFactory = geometryFactory ?? throw new ArgumentNullException(nameof(geometryFactory));
        }

        #endregion Constructors

        public override ShapeSc Create(OpenXmlCompositeElement shapeTreeSource)
        {
            Check.NotNull(shapeTreeSource, nameof(shapeTreeSource));

            P.Picture pPicture;
            if (shapeTreeSource is P.Picture treePic)
            {
                pPicture = treePic;
            }
            else
            {
                var framePic = shapeTreeSource.Descendants<P.Picture>().FirstOrDefault();
                pPicture = framePic;
            }
            if (pPicture != null)
            {
                var pBlipFill = pPicture.GetFirstChild<P.BlipFill>();
                var blipRelateId = pBlipFill?.Blip?.Embed?.Value;
                if (blipRelateId == null)
                {
                    return null;
                }
                var picture = new PictureSc(_sdkSldPart, blipRelateId);
                var spContext = _shapeContextBuilder.Build(shapeTreeSource);
                var innerTransform = _transformFactory.FromComposite(pPicture);
                var geometry = _geometryFactory.ForCompositeElement(pPicture, pPicture.ShapeProperties);
                var shape = _shapeBuilder.WithPicture(innerTransform, spContext, picture, geometry, shapeTreeSource);

                return shape;
            }

            if (Successor != null)
            {
                return Successor.Create(shapeTreeSource);
            }

            return null;
        }
    }
}
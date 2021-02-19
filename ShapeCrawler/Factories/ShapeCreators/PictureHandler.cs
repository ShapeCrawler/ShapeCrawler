using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Models;
using ShapeCrawler.Settings;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories.ShapeCreators
{
    /// <summary>
    ///     Represents a picture handler for p:pic and picture p:graphicFrame element.
    /// </summary>
    internal class PictureHandler : OpenXmlElementHandler
    {
        #region Constructors

        internal PictureHandler(ShapeContext.Builder shapeContextBuilder, LocationParser transformFactory,
            GeometryFactory geometryFactory)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
            _geometryFactory = geometryFactory ?? throw new ArgumentNullException(nameof(geometryFactory));
        }

        #endregion Constructors

        public override IShape Create(OpenXmlCompositeElement shapeTreeSource, SlideSc slide)
        {
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

                var spContext = _shapeContextBuilder.Build(shapeTreeSource);
                var innerTransform = _transformFactory.FromComposite(pPicture);
                var geometry = _geometryFactory.ForCompositeElement(pPicture, pPicture.ShapeProperties);
                var picture = new PictureSc(slide, blipRelateId, innerTransform, spContext, geometry);

                return picture;
            }

            return Successor?.Create(shapeTreeSource, slide);
        }

        #region Fields

        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly LocationParser _transformFactory;
        private readonly GeometryFactory _geometryFactory;

        #endregion Fields
    }
}
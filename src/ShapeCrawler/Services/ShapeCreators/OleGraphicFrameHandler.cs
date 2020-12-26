using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Models.Settings;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Services.Builders;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using OleObject = ShapeCrawler.Models.SlideComponents.OleObject;

namespace ShapeCrawler.Services.ShapeCreators
{
    public class OleGraphicFrameHandler : OpenXmlElementHandler
    {
        private readonly ShapeContext.Builder _shapeContextBuilder;
        //TODO: inject via DI
        private readonly LocationParser _transformFactory;
        private readonly IShapeBuilder _shapeBuilder;
        private const string Uri = "http://schemas.openxmlformats.org/presentationml/2006/ole";

        #region Constructors

        public OleGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder, LocationParser transformFactory) :
            this(shapeContextBuilder, transformFactory, new Shape.Builder())
        {
            
        }

        public OleGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder,
            LocationParser transformFactory,
            IShapeBuilder shapeBuilder)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder)); ;
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
            _shapeBuilder = shapeBuilder ?? throw new ArgumentNullException(nameof(shapeBuilder));
        }

        #endregion Constructors

        public override Shape Create(OpenXmlElement sdkElement)
        {
            Check.NotNull(sdkElement, nameof(sdkElement));

            if (sdkElement is P.GraphicFrame sdkGraphicFrame)
            {
                var grData = sdkElement.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (grData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    var spContext = _shapeContextBuilder.Build(sdkElement);
                    var innerTransform = _transformFactory.FromComposite(sdkGraphicFrame);
                    var ole = new OleObject(sdkGraphicFrame);
                    var shape = _shapeBuilder.WithOle(innerTransform, spContext, ole);

                    return shape;
                }
            }

            if (Successor != null)
            {
                return Successor.Create(sdkElement);
            }

            return null;
        }
    }
}
using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Factories
{
    internal class OleGraphicFrameHandler : OpenXmlElementHandler
    {
        private const string Uri = "http://schemas.openxmlformats.org/presentationml/2006/ole";
        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly LocationParser _transformFactory;

        #region Constructors

        internal OleGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder, LocationParser transformFactory)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
        }

        #endregion Constructors

        public override IShape Create(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide)
        {
            if (pShapeTreeChild is P.GraphicFrame pGraphicFrame)
            {
                var grData = pShapeTreeChild.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (grData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    var spContext = _shapeContextBuilder.Build(pShapeTreeChild);
                    var innerTransform = _transformFactory.FromComposite(pGraphicFrame);
                    var oleObject = new SlideOLEObject(pGraphicFrame, spContext, slide);

                    return oleObject;
                }
            }

            return Successor?.Create(pShapeTreeChild, slide);
        }
    }
}
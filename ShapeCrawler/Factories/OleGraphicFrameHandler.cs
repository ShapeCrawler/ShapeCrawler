using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class OleGraphicFrameHandler : OpenXmlElementHandler
    {
        private const string Uri = "http://schemas.openxmlformats.org/presentationml/2006/ole";
        private readonly ShapeContext.Builder shapeContextBuilder;

        #region Constructors

        internal OleGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder)
        {
            this.shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
        }

        #endregion Constructors

        public override IShape Create(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide)
        {
            if (pShapeTreeChild is P.GraphicFrame pGraphicFrame)
            {
                var grData = pShapeTreeChild.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (grData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    var spContext = shapeContextBuilder.Build(pShapeTreeChild);
                    var oleObject = new SlideOLEObject(pGraphicFrame, spContext, slide);

                    return oleObject;
                }
            }

            return Successor?.Create(pShapeTreeChild, slide);
        }
    }
}
using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Factories
{
    internal class TableGraphicFrameHandler : OpenXmlElementHandler
    {
        private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/table";
        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly LocationParser _transformFactory;

        #region Constructors

        internal TableGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder, LocationParser transformFactory)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
        }

        #endregion Constructors

        public override IShape Create(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide)
        {
            if (pShapeTreeChild is P.GraphicFrame pGraphicFrame)
            {
                A.GraphicData graphicData = pShapeTreeChild.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (graphicData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    ShapeContext spContext = _shapeContextBuilder.Build(pShapeTreeChild);
                    ILocation innerTransform = _transformFactory.FromComposite(pGraphicFrame);
                    var table = new SlideTable(pGraphicFrame, innerTransform, spContext, slide);

                    return table;
                }
            }

            return Successor?.Create(pShapeTreeChild, slide);
        }
    }
}
using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Charts;
using ShapeCrawler.Models;
using ShapeCrawler.Settings;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories.ShapeCreators
{
    internal class ChartGraphicFrameHandler : OpenXmlElementHandler
    {
        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly LocationParser _transformFactory;
        private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        #region Constructors

        internal ChartGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder, LocationParser transformFactory)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
        }

        #endregion Constructors

        public override IShape Create(OpenXmlCompositeElement shapeTreeSource, SlideSc slide)
        {
            if (shapeTreeSource is P.GraphicFrame pGraphicFrame)
            {
                A.GraphicData aGraphicData = shapeTreeSource.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (aGraphicData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    var spContext = _shapeContextBuilder.Build(shapeTreeSource);
                    var innerTransform = _transformFactory.FromComposite(pGraphicFrame);
                    var chart = new ChartSc(pGraphicFrame, slide, pGraphicFrame, innerTransform, spContext);

                    return chart;
                }
            }

            return Successor?.Create(shapeTreeSource, slide);
        }
    }
}
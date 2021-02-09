using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Charts;
using ShapeCrawler.Factories.Builders;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories.ShapeCreators
{
    public class ChartGraphicFrameHandler : OpenXmlElementHandler
    {
        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly LocationParser _transformFactory;
        private readonly IShapeBuilder _shapeBuilder;
        private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        #region Constructors

        internal ChartGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder, LocationParser transformFactory) :
            this(shapeContextBuilder, transformFactory, new ShapeSc.Builder())
        {

        }

        internal ChartGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder,
            LocationParser transformFactory,
            IShapeBuilder shapeBuilder)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
            _shapeBuilder = shapeBuilder ?? throw new ArgumentNullException(nameof(shapeBuilder));
        }

        #endregion Constructors

        public override ShapeSc Create(OpenXmlCompositeElement shapeTreeSource)
        {
            Check.NotNull(shapeTreeSource, nameof(shapeTreeSource));

            if (shapeTreeSource is P.GraphicFrame pGraphicFrame)
            {
                A.GraphicData aGraphicData = shapeTreeSource.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (aGraphicData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    var spContext = _shapeContextBuilder.Build(shapeTreeSource);
                    var innerTransform = _transformFactory.FromComposite(pGraphicFrame);
                    var chart = new ChartSc(pGraphicFrame, spContext);
                    ShapeSc shape = _shapeBuilder.WithChart(innerTransform, spContext, chart, shapeTreeSource);

                    return shape;
                }
            }

            if (Successor != null)
            {
                return Successor.Create(shapeTreeSource);
            }

            return null;
        }
    }
}
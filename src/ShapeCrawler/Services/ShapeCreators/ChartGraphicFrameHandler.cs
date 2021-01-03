using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Models.SlideComponents.Chart;
using ShapeCrawler.Services.Builders;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Services.ShapeCreators
{
    public class ChartGraphicFrameHandler : OpenXmlElementHandler
    {
        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly LocationParser _transformFactory;
        private readonly IShapeBuilder _shapeBuilder;
        private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"; //TODO: delete duplicate from GraphicFrameExtensions

        #region Constructors

        public ChartGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder, LocationParser transformFactory) :
            this(shapeContextBuilder, transformFactory, new Shape.Builder())
        {

        }

        public ChartGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder,
            LocationParser transformFactory,
            IShapeBuilder shapeBuilder)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
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
                    var chartEx = new ChartEx(sdkGraphicFrame, spContext);
                    var shape = _shapeBuilder.WithChart(innerTransform, spContext, chartEx);

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
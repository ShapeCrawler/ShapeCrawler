using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Factories.ShapeCreators
{
    internal class TableGraphicFrameHandler : OpenXmlElementHandler
    {
        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly LocationParser _transformFactory;
        private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/table";

        #region Constructors

        internal TableGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder, LocationParser transformFactory)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
        }

        #endregion Constructors

        public override IShape Create(OpenXmlCompositeElement shapeTreeSource, SlideSc slide)
        {
            Check.NotNull(shapeTreeSource, nameof(shapeTreeSource));

            if (shapeTreeSource is P.GraphicFrame pGraphicFrame)
            {
                A.GraphicData graphicData = shapeTreeSource.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (graphicData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    ShapeContext spContext = _shapeContextBuilder.Build(shapeTreeSource);
                    ILocation innerTransform = _transformFactory.FromComposite(pGraphicFrame);
                    var table = new TableSc(pGraphicFrame, innerTransform, spContext);

                    return table;
                }
            }

            return Successor?.Create(shapeTreeSource, slide);
        }
    }
}
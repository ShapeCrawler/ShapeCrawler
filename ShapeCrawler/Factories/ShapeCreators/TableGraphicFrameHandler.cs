using System;
using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using ShapeCrawler.Factories.Builders;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using ShapeCrawler.Tables;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Factories.ShapeCreators
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    public class TableGraphicFrameHandler : OpenXmlElementHandler
    {
        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly LocationParser _transformFactory;
        private readonly IShapeBuilder _shapeBuilder;
        private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/table";

        #region Constructors

        internal TableGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder, LocationParser transformFactory) :
            this(shapeContextBuilder, transformFactory, new ShapeSc.Builder())
        {
            
        }

        internal TableGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder,
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
                A.GraphicData graphicData = shapeTreeSource.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (graphicData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    ShapeContext spContext = _shapeContextBuilder.Build(shapeTreeSource);
                    ILocation innerTransform = _transformFactory.FromComposite(pGraphicFrame);
                    var table = new TableSc(pGraphicFrame);
                    ShapeSc shape = _shapeBuilder.WithTable(innerTransform, spContext, table, shapeTreeSource);

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
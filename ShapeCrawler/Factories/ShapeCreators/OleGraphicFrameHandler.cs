using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Models;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Factories.ShapeCreators
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

        public override IShape Create(OpenXmlCompositeElement shapeTreeSource, SlideSc slide)
        {
            Check.NotNull(shapeTreeSource, nameof(shapeTreeSource));

            if (shapeTreeSource is P.GraphicFrame pGraphicFrame)
            {
                var grData = shapeTreeSource.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (grData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    var spContext = _shapeContextBuilder.Build(shapeTreeSource);
                    var innerTransform = _transformFactory.FromComposite(pGraphicFrame);
                    var oleObject = new OLEObject(pGraphicFrame, innerTransform, spContext);

                    return oleObject;
                }
            }

            return Successor?.Create(shapeTreeSource, slide);
        }
    }
}
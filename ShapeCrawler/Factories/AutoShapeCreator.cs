using DocumentFormat.OpenXml;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class AutoShapeCreator : OpenXmlElementHandler
    {
        #region Constructors

        public AutoShapeCreator(ShapeContext.Builder shapeContextBuilder, LocationParser transformFactory)
        {
            _shapeContextBuilder = shapeContextBuilder;
            _transformFactory = transformFactory;
        }

        #endregion Constructors

        #region Public Methods

        public override IShape Create(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide)
        {
            if (pShapeTreeChild is P.Shape pShape)
            {
                ShapeContext shapeContext = _shapeContextBuilder.Build(pShapeTreeChild);
                ILocation innerTransform = _transformFactory.FromComposite(pShape);
                var autoShape = new SlideAutoShape(innerTransform, shapeContext, pShape, slide);

                return autoShape;
            }

            return Successor?.Create(pShapeTreeChild, slide);
        }

        #endregion Public Methods

        #region Fields

        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly LocationParser _transformFactory;

        #endregion Fields
    }
}
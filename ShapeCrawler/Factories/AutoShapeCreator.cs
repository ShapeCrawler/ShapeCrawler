using DocumentFormat.OpenXml;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class AutoShapeCreator : OpenXmlElementHandler
    {
        private readonly ShapeContext.Builder _shapeContextBuilder;

        #region Constructors

        public AutoShapeCreator(ShapeContext.Builder shapeContextBuilder)
        {
            this._shapeContextBuilder = shapeContextBuilder;
        }

        #endregion Constructors

        #region Public Methods

        public override IShape Create(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide)
        {
            if (pShapeTreeChild is P.Shape pShape)
            {
                ShapeContext shapeContext = _shapeContextBuilder.Build(pShapeTreeChild);
                var autoShape = new SlideAutoShape(shapeContext, pShape, slide);

                return autoShape;
            }

            return Successor?.Create(pShapeTreeChild, slide);
        }

        #endregion Public Methods
    }
}
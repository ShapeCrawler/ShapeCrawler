using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Factories;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a group shape on a Slide.
    /// </summary>
    internal class SlideGroupShape : SlideShape, IGroupShape
    {
        private readonly GroupShape pGroupShape;
        private readonly SCSlide slideInternal;

        public SlideGroupShape(P.GroupShape pGroupShape, SCSlide parentSlideInternal, Shape parentGroupShape)
            : base(pGroupShape, parentSlideInternal, parentGroupShape)
        {
            this.pGroupShape = pGroupShape;
            this.slideInternal = parentSlideInternal;
        }

        public IReadOnlyCollection<IShape> Shapes => this.CreateGroupedShapeCollection(); // TODO: make itlazy

        private IReadOnlyCollection<IShape> CreateGroupedShapeCollection()
        {
            SlidePart slidePart = this.slideInternal.SlidePart;
            var shapeContextBuilder = new ShapeContext.Builder(slidePart);

            var autoShapeCreator = new AutoShapeCreator();
            var oleGrFrameHandler = new OleGraphicFrameHandler(shapeContextBuilder);
            var pictureHandler = new PictureHandler();
            var chartGrFrameHandler = new ChartGraphicFrameHandler();
            var tableGrFrameHandler = new TableGraphicFrameHandler(shapeContextBuilder);

            autoShapeCreator.Successor = oleGrFrameHandler;
            oleGrFrameHandler.Successor = pictureHandler;
            pictureHandler.Successor = chartGrFrameHandler;
            chartGrFrameHandler.Successor = tableGrFrameHandler;

            List<IShape> groupedShapes = new ();
            foreach (OpenXmlCompositeElement child in this.pGroupShape.ChildElements.OfType<OpenXmlCompositeElement>())
            {
                IShape shape;
                if (child is P.GroupShape pGroupShape)
                {
                    shape = new SlideGroupShape(pGroupShape, this.slideInternal, this);
                }
                else
                {
                    shape = autoShapeCreator.Create(child, this.slideInternal, this);
                }

                if (shape != null)
                {
                    groupedShapes.Add(shape);
                }
            }

            return groupedShapes;
        }
    }
}
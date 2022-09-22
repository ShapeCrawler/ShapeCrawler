using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Factories;
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
        private readonly SCSlide slideLayoutInternal;

        public SlideGroupShape(P.GroupShape pGroupShape, SCSlide parentSlideLayoutInternal, Shape groupShape)
            : base(pGroupShape, parentSlideLayoutInternal, groupShape)
        {
            this.pGroupShape = pGroupShape;
            this.slideLayoutInternal = parentSlideLayoutInternal;
        }

        public IReadOnlyCollection<IShape> Shapes => this.CreateGroupedShapeCollection(); // TODO: make it lazy

        public SCShapeType ShapeType => SCShapeType.GroupShape;

        private IReadOnlyCollection<IShape> CreateGroupedShapeCollection()
        {
            var autoShapeCreator = new AutoShapeCreator();
            var oleGrFrameHandler = new OleGraphicFrameHandler();
            var pictureHandler = new PictureHandler();
            var chartGrFrameHandler = new ChartGraphicFrameHandler();
            var tableGrFrameHandler = new TableGraphicFrameHandler();

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
                    shape = new SlideGroupShape(pGroupShape, this.slideLayoutInternal, this);
                }
                else
                {
                    shape = autoShapeCreator.Create(child, this.slideLayoutInternal, this);
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
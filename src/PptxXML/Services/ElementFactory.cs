using ObjectEx.Utilities;
using PptxXML.Exceptions;
using PptxXML.Models.Elements;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxXML.Services
{
    /// <summary>
    /// Represents <see cref="Element"/> instance creator.
    /// </summary>
    public class ElementFactory : IElementFactory
    {
        #region Dependencies

        private readonly IGroupShapeTypeParser _groupShapeTypeParser;

        #endregion Dependencies

        #region Constructors

        public ElementFactory(IGroupShapeTypeParser groupShapeTypeParser)
        {
            _groupShapeTypeParser = groupShapeTypeParser;
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Creates instance of the <see cref="ShapeEx"/> class.
        /// </summary>
        /// <param name="ec"></param>
        /// <returns></returns>
        public Element CreateShape(ElementCandidate ec)
        {
            // Validate
            Check.NotNull(ec, nameof(ec));
            if (!(ec.CompositeElement is P.Shape xmlShape))
            {
                throw new PptxXMLException();
            }

            // Create shape
            var shape = new ShapeEx
            {
                XmlCompositeElement = xmlShape
            };

            // Add own transform properties
            var t2d = xmlShape.ShapeProperties.Transform2D;
            WithOwnTransform2d(shape, t2d);

            return shape;
        }

        /// <summary>
        /// Creates instance of the <see cref="ShapeEx"/> class.
        /// </summary>
        /// <param name="ec"></param>
        /// <returns></returns>
        public Element CreatePicture(ElementCandidate ec)
        {
            Check.NotNull(ec, nameof(ec));

            var compositeElement = ec.CompositeElement;
            if (compositeElement is P.Shape || compositeElement is P.Picture)
            {
                var t2D = compositeElement.GetFirstChild<P.ShapeProperties>().Transform2D;
                var picture = new PictureEx()
                {
                    XmlCompositeElement = compositeElement
                };
                WithOwnTransform2d(picture, t2D);

                return picture;
            }

            throw new PptxXMLException();
        }

        /// <summary>
        /// Creates instance of the <see cref="ChartEx"/> class.
        /// </summary>
        /// <param name="ec"></param>
        /// <returns></returns>
        public Element CreateChart(ElementCandidate ec)
        {
            // Validate
            Check.NotNull(ec, nameof(ec));
            if (!(ec.CompositeElement is P.GraphicFrame xmlGrFrame))
            {
                throw new PptxXMLException();
            }

            var chart = new ChartEx
            {
                XmlCompositeElement = xmlGrFrame
            };
            WithOwnTransform(chart, xmlGrFrame);

            return chart;
        }

        /// <summary>
        /// Creates instance of the <see cref="TableEx"/> class.
        /// </summary>
        /// <param name="ec"></param>
        /// <returns></returns>
        public Element CreateTable(ElementCandidate ec)
        {
            // Validate
            Check.NotNull(ec, nameof(ec));
            if (!(ec.CompositeElement is P.GraphicFrame xmlGrFrame))
            {
                throw new PptxXMLException();
            }

            var table = new TableEx
            {
                XmlCompositeElement = xmlGrFrame
            };
            WithOwnTransform(table, xmlGrFrame);

            return table;
        }

        #endregion Public Methods

        #region Private Methods

        private static void WithOwnTransform(Element e, P.GraphicFrame xmlGrFrame)
        {
            var transform = xmlGrFrame.Transform;
            e.X = transform.Offset.X.Value;
            e.Y = transform.Offset.Y.Value;
            e.Width = transform.Extents.Cx.Value;
            e.Height = transform.Extents.Cy.Value;
        }

        private static void WithOwnTransform2d(Element e, A.Transform2D t2D)
        {
            e.X = t2D.Offset.X.Value;
            e.Y = t2D.Offset.Y.Value;
            e.Width = t2D.Extents.Cx.Value;
            e.Height = t2D.Extents.Cy.Value;
        }

        #endregion Private Methods
    }
}
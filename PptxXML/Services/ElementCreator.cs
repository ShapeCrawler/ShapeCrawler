using ObjectEx.Utilities;
using PptxXML.Enums;
using PptxXML.Exceptions;
using PptxXML.Models.Elements;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Services
{
    /// <summary>
    /// Represents <see cref="Element"/> instance creator.
    /// </summary>
    public class ElementCreator : IElementCreator
    {
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
            var t2D = xmlShape.ShapeProperties.Transform2D;
            var shape = new ShapeEx(xmlShape)
            {
                Type = ElementType.Shape,
                X = t2D.Offset.X.Value,
                Y = t2D.Offset.Y.Value,
                Width = t2D.Extents.Cx.Value,
                Height = t2D.Extents.Cy.Value
            };

            return shape;
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

            var transform = xmlGrFrame.Transform;
            var chart = new ChartEx(xmlGrFrame)
            {
                X = transform.Offset.X.Value,
                Y = transform.Offset.Y.Value,
                Width = transform.Extents.Cx.Value,
                Height = transform.Extents.Cy.Value,
                Type = ElementType.Chart
            };

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

            var transform = xmlGrFrame.Transform;
            var table = new TableEx(xmlGrFrame)
            {
                X = transform.Offset.X.Value,
                Y = transform.Offset.Y.Value,
                Width = transform.Extents.Cx.Value,
                Height = transform.Extents.Cy.Value,
                Type = ElementType.Table
            };

            return table;
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
                var picture = new PictureEx(compositeElement)
                {
                    Type = ElementType.Picture,
                    X = t2D.Offset.X.Value,
                    Y = t2D.Offset.Y.Value,
                    Width = t2D.Extents.Cx.Value,
                    Height = t2D.Extents.Cy.Value
                };

                return picture;
            }

            throw new PptxXMLException();
        }
    }
}
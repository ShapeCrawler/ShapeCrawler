using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ObjectEx.Utilities;
using PptxXML.Exceptions;
using PptxXML.Models.Elements;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using PptxXML.Enums;
using PptxXML.Extensions;
using PptxXML.Models;
using PptxXML.Services.Placeholder;

namespace PptxXML.Services
{
    /// <summary>
    /// Represents factory to create elements except Group type element.
    /// </summary>
    public class ElementFactory : IElementFactory
    {
        #region Public Methods

        /// <summary>
        /// Creates a new element from root tree.
        /// </summary>
        /// <returns></returns>
        public Element CreateRootElement(ElementCandidate ec, SlidePart sldPart, Dictionary<int, PlaceholderData> phDic)
        {
            Check.NotNull(ec, nameof(ec));

            switch (ec.ElementType)
            {
                case ElementType.Shape:
                    {
                        return CreateShape(ec.CompositeElement, sldPart, phDic);
                    }
                case ElementType.Chart:
                    {
                        return CreateChart(ec);
                    }
                case ElementType.Table:
                    {
                        return CreateTable(ec);

                    }
                case ElementType.Picture:
                    {
                        return CreatePicture(ec, sldPart);
                    }
                default:
                    throw new PptxXMLException(nameof(ElementType));
            }
        }

        /// <summary>
        /// Creates a new element of group element.
        /// </summary>
        /// <returns></returns>
        public Element CreateGroupsElement(ElementCandidate ec, SlidePart sldPart)
        {
            Check.NotNull(ec, nameof(ec));

            switch (ec.ElementType)
            {
                case ElementType.Shape:
                {
                    return CreateShape(ec.CompositeElement, sldPart);
                }
                case ElementType.Chart:
                {
                    return CreateChart(ec);
                }
                case ElementType.Table:
                {
                    return CreateTable(ec);

                }
                case ElementType.Picture:
                {
                    return CreatePicture(ec, sldPart);
                }
                default:
                    throw new PptxXMLException(nameof(ElementType));
            }
        }

        #endregion Public Methods

        #region Private Methods

        private static Element CreateShape(OpenXmlCompositeElement ce, SlidePart sldPart)
        {
            // Create shape
            var shape = new ShapeEx(ce, sldPart);

            // Add own transform properties
            var t2d = ((P.Shape)ce).ShapeProperties.Transform2D;
            WithOwnTransform2d(shape, t2d);

            return shape;
        }

        private static Element CreateShape(OpenXmlCompositeElement ce, SlidePart sldPart, Dictionary<int, PlaceholderData> phDic)
        {
            // Create shape
            var shape = new ShapeEx(ce, sldPart);

            // Add own transform properties
            var t2d = ((P.Shape)ce).ShapeProperties.Transform2D;
            if (t2d != null) // not place holder
            { 
                WithOwnTransform2d(shape, t2d);
            }
            else // is placeholder
            { 
                var idx = ce.GetPlaceholderIndex();
                const string errMsg = "Something went wrong during process placeholder.";
                _ = idx ?? throw new PptxXMLException(errMsg);
                phDic.TryGetValue((int)idx, out var placeholderData);
                _ = placeholderData ?? throw new PptxXMLException(errMsg);
                shape.X = placeholderData.X;
                shape.Y = placeholderData.Y;
                shape.Width = placeholderData.Width;
                shape.Height = placeholderData.Height;
            }

            return shape;
        }

        private Element CreatePicture(ElementCandidate ec, SlidePart sldPart)
        {
            Check.NotNull(ec, nameof(ec));

            var compositeElement = ec.CompositeElement;
            if (compositeElement is P.Shape || compositeElement is P.Picture)
            {
                var t2D = compositeElement.GetFirstChild<P.ShapeProperties>().Transform2D;
                var picture = new PictureEx(sldPart, compositeElement);
                WithOwnTransform2d(picture, t2D);

                return picture;
            }

            throw new PptxXMLException();
        }

        private Element CreateChart(ElementCandidate ec)
        {
            // Validate
            Check.NotNull(ec, nameof(ec));
            if (!(ec.CompositeElement is P.GraphicFrame xmlGrFrame))
            {
                throw new PptxXMLException();
            }

            var chart = new ChartEx(xmlGrFrame);
            WithOwnTransform(chart, xmlGrFrame);

            return chart;
        }

        private Element CreateTable(ElementCandidate ec)
        {
            // Validate
            Check.NotNull(ec, nameof(ec));
            if (!(ec.CompositeElement is P.GraphicFrame xmlGrFrame))
            {
                throw new PptxXMLException();
            }

            var table = new TableEx(xmlGrFrame);
            
            WithOwnTransform(table, xmlGrFrame);

            return table;
        }

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
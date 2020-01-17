using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ObjectEx.Utilities;
using PptxXML.Exceptions;
using PptxXML.Models.Elements;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using PptxXML.Enums;
using PptxXML.Extensions;
using PptxXML.Models.Settings;
using PptxXML.Services.Builders;
using PptxXML.Services.Placeholders;

namespace PptxXML.Services
{
    /// <summary>
    /// Represents factory to create elements except Group type element.
    /// </summary>
    public class ElementFactory : IElementFactory
    {
        #region Dependencies

        private readonly IShapeExBuilder _shapeBuilder;

        #endregion Dependencies

        #region Constructors

        public ElementFactory(IShapeExBuilder shapeBuilder)
        {
            Check.NotNull(shapeBuilder, nameof(shapeBuilder));
            _shapeBuilder = shapeBuilder;
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Creates a new element from root tree.
        /// </summary>
        /// <returns></returns>
        public Element CreateRootSldElement(ElementCandidate ec, SlidePart sldPart, IPreSettings preSettings, Dictionary<int, PlaceholderEx> phDic)
        {
            Check.NotNull(ec, nameof(ec));
            var elSettings = new ElementSettings(preSettings);

            switch (ec.ElementType)
            {
                case ElementType.Shape:
                    {
                        return CreateShape(ec.CompositeElement, sldPart, elSettings, phDic);
                    }
                case ElementType.Chart:
                    {
                        return CreateChart(ec);
                    }
                case ElementType.Table:
                    {
                        return CreateTable(ec, elSettings);

                    }
                case ElementType.Picture:
                    {
                        return CreatePicture(ec, sldPart);
                    }
                case ElementType.OLEObject:
                {
                    return new OLEObject(ec.CompositeElement);
                }
                default:
                    throw new PptxXMLException(nameof(ElementType));
            }
        }

        /// <summary>
        /// Creates a new element of group element.
        /// </summary>
        /// <returns></returns>
        public Element CreateGroupsElement(ElementCandidate ec, SlidePart sldPart, IPreSettings preSettings)
        {
            Check.NotNull(ec, nameof(ec));
            var elSetting = new ElementSettings(preSettings);

            switch (ec.ElementType)
            {
                case ElementType.Shape:
                {
                    return CreateShape(ec.CompositeElement, sldPart, elSetting);
                }
                case ElementType.Chart:
                {
                    return CreateChart(ec);
                }
                case ElementType.Table:
                {
                    return CreateTable(ec, elSetting);

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

        private Element CreateShape(OpenXmlCompositeElement ce, SlidePart sldPart, ElementSettings elSettings)
        {
            // Create shape
            var shape = _shapeBuilder.Build(ce, sldPart, elSettings);

            // Add own transform properties
            var t2d = ((P.Shape)ce).ShapeProperties.Transform2D;
            WithOwnTransform2d(shape, t2d);

            return shape;
        }

        private Element CreateShape(OpenXmlCompositeElement ce, SlidePart sldPart, ElementSettings elSettings, Dictionary<int, PlaceholderEx> phDic)
        {
            ShapeEx shape;

            // Add own transform properties
            var t2d = ((P.Shape)ce).ShapeProperties.Transform2D;
            if (t2d != null)
            {
                if (ce.IsPlaceholder())
                {
                    var placeholder = GetPlaceholder(ce, phDic);
                    elSettings.Placeholder = placeholder;
                }
                shape = _shapeBuilder.Build(ce, sldPart, elSettings);
                WithOwnTransform2d(shape, t2d);
            }
            else // is placeholder obviously
            {
                var placeholder = GetPlaceholder(ce, phDic);
                elSettings.Placeholder = placeholder;

                shape = _shapeBuilder.Build(ce, sldPart, elSettings);
                shape.X = placeholder.X;
                shape.Y = placeholder.Y;
                shape.Width = placeholder.Width;
                shape.Height = placeholder.Height;
            }

            return shape;
        }

        private PlaceholderEx GetPlaceholder(OpenXmlCompositeElement ce, Dictionary<int, PlaceholderEx> phDic)
        {
            var idx = ce.GetPlaceholderIndex();
            const string errMsg = "Something went wrong during process placeholder.";
            _ = idx ?? throw new PptxXMLException(errMsg);
            phDic.TryGetValue((int)idx, out var placeholder);
            _ = placeholder ?? throw new PptxXMLException(errMsg);

            return placeholder;
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

        private Element CreateTable(ElementCandidate ec, ElementSettings elSettings)
        {
            // Validate
            Check.NotNull(ec, nameof(ec));
            if (!(ec.CompositeElement is P.GraphicFrame xmlGrFrame))
            {
                throw new PptxXMLException();
            }

            var table = new TableEx(xmlGrFrame, elSettings);
            
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
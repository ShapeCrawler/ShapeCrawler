using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideXML.Enums;
using SlideXML.Exceptions;
using SlideXML.Extensions;
using SlideXML.Models.Settings;
using SlideXML.Models.SlideComponents;
using SlideXML.Services.Builders;
using SlideXML.Services.Placeholders;
using SlideXML.Validation;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideXML.Services
{
    /// <summary>
    /// Represents slide shape factory.
    /// </summary>
    public class ElementFactory : IElementFactory
    {
        #region Fields

        private readonly IShapeBuilder _shapeBuilder;
        private readonly IPlaceholderService _phService;

        #region Dependencies

        private readonly SlidePart _sldPart;

        #endregion Dependencies

        #endregion

        #region Constructors

        public ElementFactory(SlidePart sldPart)
        {
            _sldPart = sldPart ?? throw new ArgumentNullException(nameof(sldPart));
            _shapeBuilder = new SlideElement.Builder(new BackgroundImageFactory(), new XmlGroupShapeTypeParser(), _sldPart);
            _phService = new PlaceholderService(_sldPart.SlideLayoutPart);
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Creates a new shape from candidate.
        /// </summary>
        /// <returns></returns>
        public SlideElement CreateShape(ElementCandidate ec, IPreSettings preSettings)
        {
            Check.NotNull(ec, nameof(ec));
            Check.NotNull(preSettings, nameof(preSettings));
            var elSetting = new ElementSettings(preSettings); //TODO: consider possibility use struct or WeakReference

            switch (ec.ElementType)
            {
                case ElementType.AutoShape:
                    {
                        return CreateShape(ec.CompositeElement,  elSetting);
                    }
                case ElementType.Chart:
                    {
                        return CreateChart(ec);
                    }
                case ElementType.Table:
                    {
                        var tableShape = _shapeBuilder.BuildTable((P.GraphicFrame)ec.CompositeElement, elSetting);
                        elSetting.Shape = tableShape;
                        return tableShape;

                    }
                case ElementType.Picture:
                {
                    return _shapeBuilder.BuildPicture(ec.CompositeElement);
                }
                case ElementType.OLEObject:
                {
                    return _shapeBuilder.BuildOLEObject(ec.CompositeElement);
                }
                default:
                    throw new SlideXmlException(nameof(ElementType));
            }
        }

        public SlideElement CreateGroupShape(OpenXmlCompositeElement compositeElement, IPreSettings preSettings)
        {
            return _shapeBuilder.BuildGroup(this, compositeElement, preSettings);
        }

        #endregion Public Methods

        #region Private Methods

        private SlideElement CreateShape(OpenXmlCompositeElement ce, ElementSettings elSettings)
        {
            SlideElement shape;
            // Add own transform properties
            var t2d = ((P.Shape)ce).ShapeProperties.Transform2D;
            if (t2d != null)
            {
                if (ce.IsPlaceholder())
                {
                    elSettings.Placeholder = _phService.TryGet(ce);
                }
                shape = _shapeBuilder.BuildAutoShape(ce, elSettings);
                elSettings.Shape = shape;
                WithOwnTransform2d(shape, t2d);
            }
            else // is placeholder obviously
            {
                var placeholder = _phService.TryGet(ce);
                elSettings.Placeholder = placeholder;

                shape = _shapeBuilder.BuildAutoShape(ce, elSettings);
                elSettings.Shape = shape;
                shape.X = placeholder.X;
                shape.Y = placeholder.Y;
                shape.Width = placeholder.Width;
                shape.Height = placeholder.Height;
            }

            return shape;
        }

        private SlideElement CreateChart(ElementCandidate ec)
        {
            // Validate
            Check.NotNull(ec, nameof(ec));
            if (!(ec.CompositeElement is P.GraphicFrame xmlGrFrame))
            {
                throw new SlideXmlException();
            }

            var chartShape = _shapeBuilder.BuildChart(xmlGrFrame);

            return chartShape;
        }

        private static void WithOwnTransform2d(SlideElement e, A.Transform2D t2D)
        {
            e.X = t2D.Offset.X.Value;
            e.Y = t2D.Offset.Y.Value;
            e.Width = t2D.Extents.Cx.Value;
            e.Height = t2D.Extents.Cy.Value;
        }

        #endregion Private Methods
    }
}
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Enums;
using SlideDotNet.Exceptions;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Services.Builders;
using SlideDotNet.Services.Placeholders;
using SlideDotNet.Validation;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services
{
    /// <summary>
    /// Represents slide shape factory.
    /// </summary>
    public class ElementFactory : IElementFactory
    {
        #region Fields

        private readonly IShapeBuilder _shapeBuilder;
        private readonly IPlaceholderService _phService;
        private readonly SlidePlaceholderFontService _fontService;

        #endregion

        #region Constructors

        public ElementFactory(SlidePart sldPart)
        {
            Check.NotNull(sldPart, nameof(sldPart));
            _shapeBuilder = new Shape.Builder(new BackgroundImageFactory(), new XmlGroupShapeTypeParser(), sldPart);
            _phService = new PlaceholderService(sldPart.SlideLayoutPart);
            _fontService = new SlidePlaceholderFontService(sldPart);
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Creates a new shape from candidate.
        /// </summary>
        /// <returns></returns>
        public Shape ElementFromCandidate(ElementCandidate candidate, IParents parents)
        {
            Check.NotNull(candidate, nameof(candidate));
            Check.NotNull(parents, nameof(parents));
            var elSetting = new ElementSettings(parents, _fontService, candidate.XmlElement);

            switch (candidate.ElementType)
            {
                case ElementType.AutoShape:
                    {
                        return _shapeBuilder.BuildAutoShape(candidate.XmlElement, elSetting);
                    }
                case ElementType.Chart:
                    {
                        return ChartFromXml(candidate.XmlElement);
                    }
                case ElementType.Table:
                    {
                        return TableFromXml(candidate.XmlElement, elSetting);
                    }
                case ElementType.Picture:
                    {
                        return _shapeBuilder.BuildPicture(candidate.XmlElement, elSetting);
                    }
                case ElementType.OLEObject:
                    {
                        return _shapeBuilder.BuildOleObject(candidate.XmlElement);
                    }
                default:
                    throw new SlideXmlException(nameof(ElementType));
            }
        }

        public Shape GroupFromXml(OpenXmlCompositeElement compositeElement, IParents parents)
        {
            return _shapeBuilder.BuildGroup(this, compositeElement, parents);
        }

        #endregion Public Methods

        #region Private Methods

        private Shape TableFromXml(OpenXmlCompositeElement xmlElement, ElementSettings elSetting)
        {
            var tableShape = _shapeBuilder.BuildTable((P.GraphicFrame)xmlElement, elSetting);
            elSetting.SlideElement = tableShape;

            return tableShape;
        }

        private Shape ChartFromXml(OpenXmlCompositeElement xmlElement)
        {
            var xmlGraphicFrame = (P.GraphicFrame)xmlElement;
            var chartShape = _shapeBuilder.BuildChart(xmlGraphicFrame);

            return chartShape;
        }

        #endregion Private Methods
    }
}
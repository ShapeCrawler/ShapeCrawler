using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using LogicNull.Utilities;
using SlideXML.Enums;
using SlideXML.Models.Elements;
using SlideXML.Models.Settings;
using SlideXML.Services;
using SlideXML.Services.Builders;
using SlideXML.Services.Placeholders;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Models
{
    /// <summary>
    /// Represents a slide.
    /// </summary>
    public class SlideEx
    {
        #region Fields

        private readonly SlidePart _xmlSldPart;

        private List<Element> _elements; //TODO: use capacity
        private ImageEx _backgroundImage;

        #region Dependencies

        private readonly IElementFactory _elFactory;
        private readonly IGroupShapeTypeParser _shapeTreeParser; // may be better move into _elFactory
        private readonly IGroupExBuilder _groupBuilder;
        private readonly ISlideLayoutPartParser _sldLayoutPartParser;
        private readonly IBackgroundImageFactory _bgImgFactory;
        private readonly IPreSettings _preSettings;

        #endregion Dependencies

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets elements.
        /// </summary>
        public IList<Element> Elements
        {
            get
            {
                if (_elements == null)
                {
                    InitElements();
                }

                return _elements;
            }
        }

        /// <summary>
        /// Returns a slide number in presentation.
        /// </summary>
        public int Number { get; set; } //TODO: Remove public setter somehow

        /// <summary>
        /// Returns a background image of slide.
        /// </summary>
        public ImageEx BackgroundImage
        {
            get
            {
                return _backgroundImage ??= _bgImgFactory.CreateBackgroundSlide(_xmlSldPart);
            }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initialize a new instance of the <see cref="SlideEx"/> class.
        /// </summary>
        /// TODO: use builder instead public constructor
        public SlideEx(SlidePart xmlSldPart, 
                       int sldNumber, 
                       IElementFactory elFactory, 
                       IGroupShapeTypeParser shapeTreeParser,
                       IGroupExBuilder groupBuilder,
                       ISlideLayoutPartParser sldLayoutPartParser,
                       IBackgroundImageFactory bgImgFactory,
                       IPreSettings preSettings)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));
            Check.IsPositive(sldNumber, nameof(sldNumber));
            Check.NotNull(elFactory, nameof(elFactory));
            Check.NotNull(shapeTreeParser, nameof(shapeTreeParser));
            Check.NotNull(groupBuilder, nameof(groupBuilder));
            Check.NotNull(sldLayoutPartParser, nameof(sldLayoutPartParser));
            Check.NotNull(bgImgFactory, nameof(bgImgFactory));
            _xmlSldPart = xmlSldPart;
            Number = sldNumber;
            _elFactory = elFactory;
            _shapeTreeParser = shapeTreeParser;
            _groupBuilder = groupBuilder;
            _sldLayoutPartParser = sldLayoutPartParser;
            _bgImgFactory = bgImgFactory;
            _preSettings = preSettings;
        }

        #endregion Constructors

        #region Private Methods

        private void InitElements()
        {
            // Slide
            var shTree = _xmlSldPart.Slide.CommonSlideData.ShapeTree;
            var sldCandidates = _shapeTreeParser.CreateCandidates(shTree);
            var phDic = _sldLayoutPartParser.GetPlaceholderDic(_xmlSldPart.SlideLayoutPart);
            _elements = new List<Element>(sldCandidates.Count());
            foreach (var ec in sldCandidates)
            {
                var newEl = ec.ElementType.Equals(ElementType.Group)
                    ? _groupBuilder.Build((P.GroupShape)ec.CompositeElement, _xmlSldPart, _preSettings)
                    : _elFactory.CreateRootSldElement(ec, _xmlSldPart, _preSettings, phDic);
                _elements.Add(newEl);
            }
        }

        #endregion Private Methods
    }
}
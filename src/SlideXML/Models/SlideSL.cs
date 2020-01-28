using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using LogicNull.Utilities;
using SlideXML.Enums;
using SlideXML.Models.Elements;
using SlideXML.Models.Settings;
using SlideXML.Services;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Models
{
    /// <summary>
    /// Represents a slide.
    /// </summary>
    public class SlideSL
    {
        #region Fields

        private readonly SlidePart _xmlSldPart;

        private List<ShapeSL> _shapes; //TODO: use capacity
        private ImageEx _backgroundImage;

        #region Dependencies

        private readonly IGroupShapeTypeParser _groupShapeTypeParser; // may be better move into _elFactory
        private readonly IBackgroundImageFactory _bgImgFactory;
        private readonly IPreSettings _preSettings;

        #endregion Dependencies

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets elements.
        /// </summary>
        public IList<ShapeSL> Shapes
        {
            get
            {
                if (_shapes == null)
                {
                    InitElements();
                }

                return _shapes;
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
        /// Initialize a new instance of the <see cref="SlideSL"/> class.
        /// </summary>
        /// TODO: use builder instead public constructor
        public SlideSL(SlidePart xmlSldPart, 
                       int sldNumber,
                       IGroupShapeTypeParser shapeTreeParser,
                       IBackgroundImageFactory bgImgFactory,
                       IPreSettings preSettings)
        {
            Check.NotNull(xmlSldPart, nameof(xmlSldPart));
            Check.IsPositive(sldNumber, nameof(sldNumber));
            Check.NotNull(shapeTreeParser, nameof(shapeTreeParser));
            Check.NotNull(bgImgFactory, nameof(bgImgFactory));
            _xmlSldPart = xmlSldPart;
            Number = sldNumber;
            _groupShapeTypeParser = shapeTreeParser;
            _bgImgFactory = bgImgFactory;
            _preSettings = preSettings;
        }

        #endregion Constructors

        #region Private Methods

        private void InitElements()
        {
            // Slide
            var elFactory = new ElementFactory(_xmlSldPart);
            var sldCandidates = _groupShapeTypeParser.CreateCandidates(_xmlSldPart.Slide.CommonSlideData.ShapeTree);
            _shapes = new List<ShapeSL>(sldCandidates.Count());
            foreach (var candidate in sldCandidates)
            {
                ShapeSL newShape;
                if (candidate.ElementType == ShapeType.Group)
                {
                    newShape = elFactory.CreateGroupShape(candidate.CompositeElement, _preSettings);
                }
                else
                {
                    newShape = elFactory.CreateShape(candidate, _preSettings);
                }
                _shapes.Add(newShape);
            }
        }

        #endregion Private Methods
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Enums;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Services;
using SlideDotNet.Validation;

// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Models
{
    /// <summary>
    /// Represents a slide.
    /// </summary>
    public class Slide
    {
        #region Fields

        private readonly SlidePart _xmlSldPart;

        private List<Shape> _elements;
        private ImageEx _backgroundImage;

        #region Dependencies

        private readonly IXmlGroupShapeTypeParser _groupShapeTypeParser; // may be better move into _elFactory
        private readonly IBackgroundImageFactory _bgImgFactory;
        private readonly IParents _parents;

        #endregion Dependencies

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets slide elements.
        /// </summary>
        public IList<Shape> Elements
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
        /// Initialize a new instance of the <see cref="Slide"/> class.
        /// </summary>
        /// TODO: use builder instead public constructor
        public Slide(SlidePart xmlSldPart, int sldNumber, IParents parents)
        {
            Check.IsPositive(sldNumber, nameof(sldNumber));
            Number = sldNumber;
            _xmlSldPart = xmlSldPart ?? throw new ArgumentNullException(nameof(xmlSldPart));
            _groupShapeTypeParser = new XmlGroupShapeTypeParser();
            _bgImgFactory = new BackgroundImageFactory();
            _parents = parents ?? throw new ArgumentNullException(nameof(parents));
        }

        #endregion Constructors

        #region Private Methods

        private void InitElements()
        {
            var elCandidates = _groupShapeTypeParser.CreateElementCandidates(_xmlSldPart.Slide.CommonSlideData.ShapeTree);
            _elements = new List<Shape>(elCandidates.Count());
            var elFactory = new ElementFactory(_xmlSldPart);
            foreach (var candidate in elCandidates)
            {
                Shape newElement;
                if (candidate.ElementType == ElementType.Group)
                {
                    newElement = elFactory.GroupFromXml(candidate.XmlElement, _parents);
                }
                else
                {
                    newElement = elFactory.ElementFromCandidate(candidate, _parents);
                }
                _elements.Add(newElement);
            }
        }

        #endregion Private Methods
    }
}
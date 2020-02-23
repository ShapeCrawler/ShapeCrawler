using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Services;

// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Models
{
    /// <summary>
    /// Represents a slide.
    /// </summary>
    public class Slide
    {
        #region Fields

        private ImageEx _backgroundImg;
        private readonly IPreSettings _preSettings;
        private readonly SlidePart _xmlSldPart;
        private readonly Lazy<List<ShapeEx>> _shapes;
        private readonly BackgroundImageFactory _backgroundImageFactory = new BackgroundImageFactory(); //TODO: [DI]
        private readonly SlideNumber _sldNumEntity;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets slide elements.
        /// </summary>
        public IList<ShapeEx> Shapes => _shapes.Value;

        /// <summary>
        /// Returns a slide number in presentation.
        /// </summary>
        public int Number => _sldNumEntity.Number;

        /// <summary>
        /// Returns a background image of the slide. Returns null if slide does not have background image.
        /// </summary>
        public ImageEx BackgroundImage
        {
            get
            {
                return _backgroundImg ??= _backgroundImageFactory.FromXmlSlide(_xmlSldPart);
            }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initialize a new instance of the <see cref="Slide"/> class.
        /// </summary>
        /// TODO: use builder instead public constructor
        public Slide(SlidePart xmlSldPart, SlideNumber sldNum, IPreSettings preSettings)
        {
            _xmlSldPart = xmlSldPart ?? throw new ArgumentNullException(nameof(xmlSldPart));
            _sldNumEntity = sldNum ?? throw new ArgumentNullException(nameof(SlideNumber));
            _preSettings = preSettings ?? throw new ArgumentNullException(nameof(preSettings));
            _shapes = new Lazy<List<ShapeEx>>(GetShapes);
        }

        #endregion Constructors

        #region Private Methods

        private List<ShapeEx> GetShapes()
        {
            var shapeFactory = new ShapeFactory(_xmlSldPart, _preSettings);
            return shapeFactory.CreateShapesCollection(_xmlSldPart.Slide.CommonSlideData.ShapeTree).ToList();
        }

        #endregion Private Methods
    }
}
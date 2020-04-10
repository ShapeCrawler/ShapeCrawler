using System;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models.Settings;
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

        private readonly Lazy<ImageEx> _backgroundImage;
        private readonly Lazy<ShapeCollection> _shapes;
        private readonly IPreSettings _preSettings;
        private readonly SlidePart _sdkSldPart;
        private readonly SlideNumber _sldNumEntity;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Returns a slide shapes.
        /// </summary>
        public ShapeCollection Shapes => _shapes.Value;

        /// <summary>
        /// Returns a slide number in presentation.
        /// </summary>
        public int Number => _sldNumEntity.Number;

        /// <summary>
        /// Returns a background image of the slide. Returns <c>null</c>if slide does not have background image.
        /// </summary>
        public ImageEx BackgroundImage => _backgroundImage.Value;

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Slide"/> class.
        /// </summary>
        public Slide(SlidePart sdkSldPart, SlideNumber sldNum, IPreSettings preSettings)
        {
            _sdkSldPart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _sldNumEntity = sldNum ?? throw new ArgumentNullException(nameof(SlideNumber));
            _preSettings = preSettings ?? throw new ArgumentNullException(nameof(preSettings));

            _shapes = new Lazy<ShapeCollection>(GetShapeCollection);
            _backgroundImage = new Lazy<ImageEx>(TryGetBackground);
        }

        #endregion Constructors

        #region Private Methods

        private ShapeCollection GetShapeCollection()
        {
            var shapeCollection = new ShapeCollection(_sdkSldPart, _preSettings);
            return shapeCollection;
        }

        private ImageEx TryGetBackground()
        {
            var backgroundImageFactory = new ImageExFactory();
            return backgroundImageFactory.TryFromXmlSlide(_sdkSldPart);
        }

        #endregion Private Methods
    }
}
using System;
using System.Drawing;
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

        public void SaveScheme(string filePath)
        {
            var sldWidthEmu = _preSettings.SlideSize.Value.Width;
            var sldHeightEmu = _preSettings.SlideSize.Value.Height;
            var sldWidthPx = sldWidthEmu / 10000;
            var sldHeightPx = sldHeightEmu / 10000;

            using var bitmap = new Bitmap(sldWidthPx+50, sldHeightPx+50);
            var graphics = Graphics.FromImage(bitmap);

            using var blackPen = new Pen(Color.Black, 3); // create pen.
            var rect = new Rectangle(253, 500, 189, 62); // create rectangle. (x,y,width, height)
            var sldRectangle = new Rectangle(10, 10, sldWidthPx, sldHeightPx); // create rectangle. (x,y,width, height)

            var rect2 = new Rectangle(651, 194, 189, 81);

            graphics.DrawRectangle(blackPen, sldRectangle);

            //graphics.DrawRectangle(blackPen, rect);

            bitmap.Save(@"d:\1\test.png");
        }

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
using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Models.Settings;
using ShapeCrawler.Services.Drawing;
using ShapeCrawler.Statics;
using SlideDotNet.Models;

// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Models
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
        private readonly ISlideSchemeService _schemeService;
        private readonly SlidePart _sdkSldPart;
        private readonly SlideNumber _sldNumEntity;
        private readonly Lazy<CustomXmlPart> _sldXmlPart;

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
        public ImageEx Background => _backgroundImage.Value;

        public string CustomData
        {
            get => GetCustomData();
            set => SetCustomData(value);
        }

        public bool Hidden => _sdkSldPart.Slide.Show != null && _sdkSldPart.Slide.Show.Value == false;

        #endregion Properties

        #region Constructors

        public Slide(SlidePart sdkSldPart, SlideNumber sldNum, IPreSettings preSettings) :
            this(sdkSldPart, sldNum, preSettings, new SlideSchemeService())
        {

        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Slide"/> class.
        /// </summary>
        public Slide(SlidePart sdkSldPart, SlideNumber sldNum, IPreSettings preSettings, ISlideSchemeService schemeService)
        {
            _sdkSldPart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _sldNumEntity = sldNum ?? throw new ArgumentNullException(nameof(sldNum));
            _preSettings = preSettings ?? throw new ArgumentNullException(nameof(preSettings));
            _schemeService = schemeService ?? throw new ArgumentNullException(nameof(schemeService));
            _shapes = new Lazy<ShapeCollection>(GetShapeCollection);
            _backgroundImage = new Lazy<ImageEx>(TryGetBackground);
            _sldXmlPart = new Lazy<CustomXmlPart>(GetSldCustomXmlPart);
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Saves slide scheme in PNG file.
        /// </summary>
        /// <param name="filePath"></param>
        public void SaveScheme(string filePath)
        {
            var sldSize = _preSettings.SlideSize.Value;
            _schemeService.SaveScheme(_shapes.Value, sldSize.Width, sldSize.Height, filePath);
        }

        /// <summary>
        /// Saves slide scheme in stream.
        /// </summary>
        /// <param name="stream"></param>
        public void SaveScheme(Stream stream)
        {
            var sldSize = _preSettings.SlideSize.Value;
            _schemeService.SaveScheme(_shapes.Value, sldSize.Width, sldSize.Height, stream);
        }

        #endregion Public Methods

        #region Private Methods

        private ShapeCollection GetShapeCollection()
        {
            var shapeCollection = new ShapeCollection(_sdkSldPart, _preSettings);
            return shapeCollection;
        }

        private ImageEx TryGetBackground()
        {
            var backgroundImageFactory = new ImageExFactory();
            return backgroundImageFactory.TryFromSdkSlide(_sdkSldPart);
        }

        private string GetCustomData()
        {
            using var sr = new StreamReader(_sldXmlPart.Value.GetStream());
            var raw = sr.ReadToEnd();
            if (raw.Length == 0)
            {
                return null;
            }

            return raw.Substring(ConstantStrings.CustomDataElementName.Length);
        }

        private void SetCustomData(string value)
        {
            var sldXmlPartStream = _sldXmlPart.Value.GetStream();
            using var streamWriter = new StreamWriter(sldXmlPartStream);
            streamWriter.Write($"{ConstantStrings.CustomDataElementName}{value}");
        }

        private CustomXmlPart GetSldCustomXmlPart()
        {
            foreach (var customXmlPart in _sdkSldPart.CustomXmlParts)
            {
                using var streamReader = new StreamReader(customXmlPart.GetStream());
                string customXmlPartText = streamReader.ReadToEnd();
                if (customXmlPartText.StartsWith(ConstantStrings.CustomDataElementName, StringComparison.CurrentCulture))
                {
                    return customXmlPart;
                }
            }

            return _sdkSldPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        }

        #endregion Private Methods

        public void Hide()
        {
            if (_sdkSldPart.Slide.Show == null)
            {
                var showAttribute = new OpenXmlAttribute("show", "", "0");
                _sdkSldPart.Slide.SetAttribute(showAttribute);
            }
            else
            {
                _sdkSldPart.Slide.Show = false;
            }
        }
    }
}
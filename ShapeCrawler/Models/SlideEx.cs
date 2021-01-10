using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Factories.Drawing;
using ShapeCrawler.Settings;
using ShapeCrawler.Statics;
using SlideDotNet.Models;

// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Models
{
    /// <summary>
    /// Represents a slide.
    /// </summary>
    public class SlideEx
    {
        #region Fields

        private readonly Lazy<ImageEx> _backgroundImage;
        private readonly Lazy<ShapesCollection> _shapes;
        private readonly IPresentationData _preSettings;
        private readonly SlidePart _sdkSldPart;
        private readonly SlideNumber _sldNumEntity;
        private Lazy<CustomXmlPart> _customXmlPart;

        public PresentationEx PresentationEx { get; }

        #endregion Fields

        #region Properties

        /// <summary>
        /// Returns a slide shapes.
        /// </summary>
        public ShapesCollection Shapes => _shapes.Value;

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

        public SlideEx(SlidePart sdkSldPart, SlideNumber sldNum, IPresentationData preSettings, PresentationEx presentationEx) :
            this(sdkSldPart, sldNum, preSettings, new SlideSchemeService(), presentationEx)
        {

        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SlideEx"/> class.
        /// </summary>
        public SlideEx(
            SlidePart sdkSldPart, 
            SlideNumber sldNum, 
            IPresentationData preSettings, 
            SlideSchemeService schemeService, 
            PresentationEx presentationEx)
        {
            _sdkSldPart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _sldNumEntity = sldNum ?? throw new ArgumentNullException(nameof(sldNum));
            _preSettings = preSettings ?? throw new ArgumentNullException(nameof(preSettings));
            _shapes = new Lazy<ShapesCollection>(GetShapeCollection);
            _backgroundImage = new Lazy<ImageEx>(TryGetBackground);
            _customXmlPart = new Lazy<CustomXmlPart>(GetSldCustomXmlPart);
            PresentationEx = presentationEx;
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
            SlideSchemeService.SaveScheme(_shapes.Value, sldSize.Width, sldSize.Height, filePath);
        }

        /// <summary>
        /// Saves slide scheme in stream.
        /// </summary>
        /// <param name="stream"></param>
        public void SaveScheme(Stream stream)
        {
            var sldSize = _preSettings.SlideSize.Value;
            SlideSchemeService.SaveScheme(_shapes.Value, sldSize.Width, sldSize.Height, stream);
        }

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

        #endregion Public Methods

        #region Private Methods

        private ShapesCollection GetShapeCollection()
        {
            var shapeCollection = new ShapesCollection(_sdkSldPart, _preSettings, this);
            return shapeCollection;
        }

        private ImageEx TryGetBackground()
        {
            var backgroundImageFactory = new ImageExFactory();
            return backgroundImageFactory.TryFromSdkSlide(_sdkSldPart);
        }

        private string GetCustomData()
        {
            if (_customXmlPart.Value == null)
            {
                return null;
            }

            var customXmlPartStream = _customXmlPart.Value.GetStream();
            using var customXmlStreamReader = new StreamReader(customXmlPartStream);
            var raw = customXmlStreamReader.ReadToEnd();

            return raw.Substring(ConstantStrings.CustomDataElementName.Length);
        }

        private void SetCustomData(string value)
        {
            Stream customXmlPartStream;
            if (_customXmlPart.Value == null)
            {
                var newSlideCustomXmlPart = _sdkSldPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                customXmlPartStream = newSlideCustomXmlPart.GetStream();
#if NETSTANDARD2_0
                _customXmlPart = new Lazy<CustomXmlPart>(()=>newSlideCustomXmlPart);
#else
                _customXmlPart = new Lazy<CustomXmlPart>(newSlideCustomXmlPart);
#endif
            }
            else
            {
                customXmlPartStream = _customXmlPart.Value.GetStream();
            }
            using var customXmlStreamReader = new StreamWriter(customXmlPartStream);
            customXmlStreamReader.Write($"{ConstantStrings.CustomDataElementName}{value}");
        }

        private CustomXmlPart GetSldCustomXmlPart()
        {
            foreach (var customXmlPart in _sdkSldPart.CustomXmlParts)
            {
                using var customXmlPartStream = new StreamReader(customXmlPart.GetStream());
                string customXmlPartText = customXmlPartStream.ReadToEnd();
                if (customXmlPartText.StartsWith(ConstantStrings.CustomDataElementName, StringComparison.CurrentCulture))
                {
                    return customXmlPart;
                }
            }

            return null;
        }

#endregion Private Methods
    }
}
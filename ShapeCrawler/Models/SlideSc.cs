using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Factories.Drawing;
using ShapeCrawler.Models;
using ShapeCrawler.Settings;
using ShapeCrawler.SlideMaster;
using ShapeCrawler.Statics;
using SkiaSharp;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <summary>
    /// Represents a slide.
    /// </summary>
    public class SlideSc
    {
        #region Fields

        private readonly Lazy<ImageSc> _backgroundImage;
        private readonly Lazy<ShapesCollection> _shapes;
        private readonly PresentationData _preSettings;
        private readonly SlidePart _slidePart;
        private readonly SlideNumber _sldNumEntity;
        private Lazy<CustomXmlPart> _customXmlPart;

        public PresentationSc PresentationEx { get; }

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
        public ImageSc Background => _backgroundImage.Value;

        public string CustomData
        {
            get => GetCustomData();
            set => SetCustomData(value);
        }

        public bool Hidden => _slidePart.Slide.Show != null && _slidePart.Slide.Show.Value == false;
        public SlideLayoutSc Layout => GetSlideLayout();

        private SlideLayoutSc GetSlideLayout()
        {
            throw new NotImplementedException();
        }

        #endregion Properties

        #region Constructors

        public SlideSc(SlidePart sdkSldPart, SlideNumber sldNum, PresentationData preSettings, PresentationSc presentationEx) :
            this(sdkSldPart, sldNum, preSettings, new SlideSchemeService(), presentationEx)
        {

        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SlideSc"/> class.
        /// </summary>
        public SlideSc(
            SlidePart sdkSldPart, 
            SlideNumber sldNum, 
            PresentationData preSettings, 
            SlideSchemeService schemeService, 
            PresentationSc presentationEx)
        {
            _slidePart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _sldNumEntity = sldNum ?? throw new ArgumentNullException(nameof(sldNum));
            _preSettings = preSettings ?? throw new ArgumentNullException(nameof(preSettings));
            _shapes = new Lazy<ShapesCollection>(GetShapesCollection);
            _backgroundImage = new Lazy<ImageSc>(TryGetBackground);
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
#if DEBUG
        public void SaveImage(string filePath)
        {
            ShapesCollection shapes = Shapes;

            SKImageInfo imageInfo = new SKImageInfo(500, 600);
            using SKSurface surface = SKSurface.Create(imageInfo);
            SKCanvas canvas = surface.Canvas;

            canvas.Clear(SKColors.Red);

            using SKPaint paint = new SKPaint
            {
                Color = SKColors.Blue,
                IsAntialias = true,
                StrokeWidth = 15,
                Style = SKPaintStyle.Stroke
            };
            canvas.DrawCircle(70,70,50, paint);

            using SKPaint textPaint = new SKPaint();
            textPaint.Color = SKColors.Green;
            textPaint.IsAntialias = true;
            textPaint.TextSize = 48;

            using SKImage image = surface.Snapshot();
            using SKData data = image.Encode(SKEncodedImageFormat.Png, 100);
            File.WriteAllBytes(filePath, data.ToArray());
        }
#endif

        public void Hide()
        {
            if (_slidePart.Slide.Show == null)
            {
                var showAttribute = new OpenXmlAttribute("show", "", "0");
                _slidePart.Slide.SetAttribute(showAttribute);
            }
            else
            {
                _slidePart.Slide.Show = false;
            }
        }

        #endregion Public Methods

        #region Private Methods

        private ShapesCollection GetShapesCollection()
        {
            return ShapesCollection.CreateForUserSlide(_slidePart, _preSettings, this);
        }

        private ImageSc TryGetBackground()
        {
            var backgroundImageFactory = new ImageExFactory();
            return backgroundImageFactory.TryFromSdkSlide(_slidePart);
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
                var newSlideCustomXmlPart = _slidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
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
            foreach (var customXmlPart in _slidePart.CustomXmlParts)
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
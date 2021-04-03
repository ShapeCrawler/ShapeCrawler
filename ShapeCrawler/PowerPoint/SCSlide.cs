using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Drawing;
using ShapeCrawler.Factories;
using ShapeCrawler.Models;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMaster;
using ShapeCrawler.Statics;
using SkiaSharp;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a slide.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class SCSlide : ISlide // TODO: make it internal
    {
        #region Fields

        private readonly Lazy<SCImage> _backgroundImage;
        protected ResettableLazy<ShapeCollection> _shapes { get; set; }
        private readonly SlideNumber _sldNumEntity;
        private Lazy<CustomXmlPart> _customXmlPart;

        internal SCPresentation Presentation { get; }
        internal SlidePart SlidePart { get; }
        internal SCSlideLayout SlideLayout => Presentation.SlideMasters.GetSlideLayoutBySlide(this);

        #endregion Fields

        #region Public Properties

        /// <summary>
        ///     Returns a slide shapes.
        /// </summary>
        public ShapeCollection Shapes => _shapes.Value;

        /// <summary>
        ///     Returns a slide number in presentation.
        /// </summary>
        public int Number => _sldNumEntity.Number;

        /// <summary>
        ///     Returns a background image of the slide. Returns <c>null</c>if slide does not have background image.
        /// </summary>
        public SCImage Background => _backgroundImage.Value;

        public string CustomData
        {
            get => GetCustomData();
            set => SetCustomData(value);
        }

        public bool Hidden => SlidePart.Slide.Show != null && SlidePart.Slide.Show.Value == false;

        #endregion Public Properties

        #region Constructors

        /// <summary>
        ///     Initializes a new instance of the <see cref="SCSlide" /> class.
        /// </summary>
        internal SCSlide(SCPresentation presentation,
            SlidePart slidePart,
            SlideNumber sldNum)
        {
            Presentation = presentation;
            SlidePart = slidePart;
            _sldNumEntity = sldNum;
            _shapes = new ResettableLazy<ShapeCollection>(() => ShapeCollection.CreateForSlide(SlidePart, this));
            _backgroundImage = new Lazy<SCImage>(TryGetBackground);
            _customXmlPart = new Lazy<CustomXmlPart>(GetSldCustomXmlPart);
        }

        protected SCSlide(SCPresentation presentation, SlidePart slidePart)
        {
            Presentation = presentation;
            SlidePart = slidePart;
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        ///     Saves slide scheme in PNG file.
        /// </summary>
        /// <param name="filePath"></param>
        public void SaveScheme(string filePath)
        {
            SlideSchemeService.SaveScheme(_shapes.Value, Presentation.SlideWidth, Presentation.SlideHeight, filePath);
        }

        /// <summary>
        ///     Saves slide scheme in stream.
        /// </summary>
        /// <param name="stream"></param>
        public void SaveScheme(Stream stream)
        {
            SlideSchemeService.SaveScheme(_shapes.Value, Presentation.SlideWidth, Presentation.SlideHeight, stream);
        }
#if DEBUG
        public void SaveImage(string filePath)
        {
            ShapeCollection shapes = Shapes;

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
            canvas.DrawCircle(70, 70, 50, paint);

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
            if (SlidePart.Slide.Show == null)
            {
                var showAttribute = new OpenXmlAttribute("show", "", "0");
                SlidePart.Slide.SetAttribute(showAttribute);
            }
            else
            {
                SlidePart.Slide.Show = false;
            }
        }

        #endregion Public Methods

        #region Private Methods

        private SCImage TryGetBackground()
        {
            var backgroundImageFactory = new ImageExFactory();
            return backgroundImageFactory.TryFromSdkSlide(SlidePart);
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
#if NET5_0
            return raw[ConstantStrings.CustomDataElementName.Length..];
#else
            return raw.Substring(ConstantStrings.CustomDataElementName.Length);
#endif
        }

        private void SetCustomData(string value)
        {
            Stream customXmlPartStream;
            if (_customXmlPart.Value == null)
            {
                var newSlideCustomXmlPart = SlidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                customXmlPartStream = newSlideCustomXmlPart.GetStream();
#if NET461
                _customXmlPart = new Lazy<CustomXmlPart>(() => newSlideCustomXmlPart);
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
            foreach (var customXmlPart in SlidePart.CustomXmlParts)
            {
                using var customXmlPartStream = new StreamReader(customXmlPart.GetStream());
                string customXmlPartText = customXmlPartStream.ReadToEnd();
                if (customXmlPartText.StartsWith(ConstantStrings.CustomDataElementName,
                    StringComparison.CurrentCulture))
                {
                    return customXmlPart;
                }
            }

            return null;
        }

        #endregion Private Methods
    }
}
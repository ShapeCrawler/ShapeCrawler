using System;
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
    internal class SCSlide : ISlide, IRemovable // TODO: make it internal
    {
        private readonly Lazy<SCImage> backgroundImage;
        private Lazy<CustomXmlPart> customXmlPart;

        /// <summary>
        ///     Initializes a new instance of the <see cref="SCSlide" /> class.
        /// </summary>
        internal SCSlide(
            SCPresentation presentation,
            SlidePart slidePart,
            int slideNumber)
        {
            this.ParentPresentation = presentation;
            this.SlidePart = slidePart;
            this.Number = slideNumber;
            this._shapes = new ResettableLazy<ShapeCollection>(() => ShapeCollection.CreateForSlide(this.SlidePart, this));
            this.backgroundImage = new Lazy<SCImage>(this.TryGetBackground);
            this.customXmlPart = new Lazy<CustomXmlPart>(this.GetSldCustomXmlPart);
        }

        protected ResettableLazy<ShapeCollection> _shapes { get; }

        internal SCPresentation ParentPresentation { get; }

        internal SlidePart SlidePart { get; }

        #region Public Properties

        public SCSlideLayout Layout => this.ParentPresentation.SlideMasters.GetSlideLayoutBySlide(this);

        /// <summary>
        ///     Returns a slide shapes.
        /// </summary>
        public ShapeCollection Shapes => _shapes.Value;

        /// <summary>
        ///     Gets a slide number in presentation.
        /// </summary>
        public int Number { get; }

        /// <summary>
        ///     Returns a background image of the slide. Returns <c>null</c>if slide does not have background image.
        /// </summary>
        public SCImage Background => backgroundImage.Value;

        public string CustomData
        {
            get => this.GetCustomData();
            set => this.SetCustomData(value);
        }

        public bool Hidden => SlidePart.Slide.Show != null && SlidePart.Slide.Show.Value == false;

        public bool IsRemoved { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        ///     Saves slide scheme in PNG file.
        /// </summary>
        public void SaveScheme(string filePath)
        {
            SlideSchemeService.SaveScheme(_shapes.Value, ParentPresentation.SlideWidth, ParentPresentation.SlideHeight, filePath);
        }

        /// <summary>
        ///     Saves slide scheme in stream.
        /// </summary>
        public void SaveScheme(Stream stream)
        {
            SlideSchemeService.SaveScheme(_shapes.Value, ParentPresentation.SlideWidth, ParentPresentation.SlideHeight, stream);
        }

#if DEBUG
        public void SaveImage(string filePath)
        {
            ShapeCollection shapes = Shapes;

            SKImageInfo imageInfo = new(500, 600);
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
                var showAttribute = new OpenXmlAttribute("show", string.Empty, "0");
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
            var backgroundImageFactory = new SCImageFactory();
            return backgroundImageFactory.FromSlidePart(SlidePart);
        }

        private string GetCustomData()
        {
            if (customXmlPart.Value == null)
            {
                return null;
            }

            var customXmlPartStream = customXmlPart.Value.GetStream();
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
            if (customXmlPart.Value == null)
            {
                var newSlideCustomXmlPart = SlidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                customXmlPartStream = newSlideCustomXmlPart.GetStream();
#if NETSTANDARD2_0
                customXmlPart = new Lazy<CustomXmlPart>(() => newSlideCustomXmlPart);
#else
                customXmlPart = new Lazy<CustomXmlPart>(newSlideCustomXmlPart);
#endif
            }
            else
            {
                customXmlPartStream = customXmlPart.Value.GetStream();
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
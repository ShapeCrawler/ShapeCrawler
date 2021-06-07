using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Factories;
using ShapeCrawler.Models;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Statics;
using SkiaSharp;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a slide.
    /// </summary>
    internal class SCSlide : ISlide
    {
        private readonly Lazy<SCImage> backgroundImage;
        private Lazy<CustomXmlPart> customXmlPart;

        /// <summary>
        ///     Initializes a new instance of the <see cref="SCSlide" /> class.
        /// </summary>
        internal SCSlide(
            SCPresentation parentPresentation,
            SlidePart slidePart,
            int slideNumber)
        {
            this.ParentPresentation = parentPresentation;
            this.SlidePart = slidePart;
            this.Number = slideNumber;
            this._shapes = new ResettableLazy<ShapeCollection>(() => ShapeCollection.CreateForSlide(this.SlidePart, this));
            this.backgroundImage = new Lazy<SCImage>(() => SCImage.GetSlideBackgroundImageOrDefault(this));
            this.customXmlPart = new Lazy<CustomXmlPart>(this.GetSldCustomXmlPart);
        }

        #region Public Properties

        public ISlideLayout ParentSlideLayout => ((SlideMasterCollection)this.ParentPresentation.SlideMasters).GetSlideLayoutBySlide(this);

        public IShapeCollection Shapes => this._shapes.Value;

        public int Number { get; }

        public SCImage Background => this.backgroundImage.Value;

        public string CustomData
        {
            get => this.GetCustomData();
            set => this.SetCustomData(value);
        }

        public bool Hidden => this.SlidePart.Slide.Show != null && this.SlidePart.Slide.Show.Value == false;

        #endregion Public Properties

        public bool IsRemoved { get; set; }

        internal SlidePart SlidePart { get; }

        protected ResettableLazy<ShapeCollection> _shapes { get; }

        public SCPresentation ParentPresentation { get; }

        /// <summary>
        ///     Saves slide scheme in PNG file.
        /// </summary>
        public void SaveScheme(string filePath)
        {
            SlideSchemeService.SaveScheme(this._shapes.Value, this.ParentPresentation.SlideWidth, this.ParentPresentation.SlideHeight, filePath);
        }

        /// <summary>
        ///     Saves slide scheme in stream.
        /// </summary>
        public void SaveScheme(Stream stream)
        {
            SlideSchemeService.SaveScheme(this._shapes.Value, this.ParentPresentation.SlideWidth, this.ParentPresentation.SlideHeight, stream);
        }

#if DEBUG
        public void SaveImage(string filePath)
        {
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
            if (this.SlidePart.Slide.Show == null)
            {
                var showAttribute = new OpenXmlAttribute("show", string.Empty, "0");
                this.SlidePart.Slide.SetAttribute(showAttribute);
            }
            else
            {
                this.SlidePart.Slide.Show = false;
            }
        }

        #region Private Methods

        private string GetCustomData()
        {
            if (this.customXmlPart.Value == null)
            {
                return null;
            }

            var customXmlPartStream = this.customXmlPart.Value.GetStream();
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
            if (this.customXmlPart.Value == null)
            {
                var newSlideCustomXmlPart = this.SlidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                customXmlPartStream = newSlideCustomXmlPart.GetStream();
#if NETSTANDARD2_0
                customXmlPart = new Lazy<CustomXmlPart>(() => newSlideCustomXmlPart);
#else
                this.customXmlPart = new Lazy<CustomXmlPart>(newSlideCustomXmlPart);
#endif
            }
            else
            {
                customXmlPartStream = this.customXmlPart.Value.GetStream();
            }

            using var customXmlStreamReader = new StreamWriter(customXmlPartStream);
            customXmlStreamReader.Write($"{ConstantStrings.CustomDataElementName}{value}");
        }

        private CustomXmlPart GetSldCustomXmlPart()
        {
            foreach (CustomXmlPart customXmlPart in this.SlidePart.CustomXmlParts)
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

        public void ThrowIfRemoved()
        {
            if (this.IsRemoved)
            {
                throw new ElementIsRemovedException("Slide was removed");
            }
            else
            {
                this.ParentPresentation.ThrowIfClosed();
            }
        }

        #endregion Private Methods
    }
}
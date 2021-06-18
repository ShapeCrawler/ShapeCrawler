﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Factories;
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
            SlidePart slidePart)
        {
            this.ParentPresentation = parentPresentation;
            this.SlidePart = slidePart;
            this._shapes = new ResettableLazy<ShapeCollection>(() => ShapeCollection.CreateForSlide(this.SlidePart, this));
            this.backgroundImage = new Lazy<SCImage>(() => SCImage.GetSlideBackgroundImageOrDefault(this));
            this.customXmlPart = new Lazy<CustomXmlPart>(this.GetSldCustomXmlPart);
        }

        #region Public Properties

        public ISlideLayout ParentSlideLayout => ((SlideMasterCollection)this.ParentPresentation.SlideMasters).GetSlideLayoutBySlide(this);

        public IShapeCollection Shapes => this._shapes.Value;

        public int Number
        {
            get => this.GetNumber();
            set => this.SetNumber(value);
        }

        private int GetNumber()
        {
            string currentSlidePartId =
                this.ParentPresentation.PresentationDocument.PresentationPart.GetIdOfPart(this.SlidePart);
            List<SlideId> slideIdList = this.ParentPresentation.PresentationDocument.PresentationPart.Presentation
                .SlideIdList.ChildElements.OfType<SlideId>().ToList();
            for (int i = 0; i < slideIdList.Count; i++)
            {
                if (slideIdList[i].RelationshipId == currentSlidePartId)
                {
                    return i + 1;
                }
            }

            throw new ShapeCrawlerException("An error occurred while parsing slide number.");
        }

        private void SetNumber(int newSlideNumber)
        {
            int from = this.Number - 1;
            int to = newSlideNumber - 1;

            if (to < 0 || from >= this.ParentPresentation.Slides.Count || to == from)
            {
                throw new ArgumentOutOfRangeException(nameof(to));
            }

            PresentationPart presentationPart = this.ParentPresentation.PresentationDocument.PresentationPart;

            Presentation presentation = presentationPart.Presentation;
            SlideIdList slideIdList = presentation.SlideIdList;

            // Get the slide ID of the source slide.
            SlideId sourceSlide = slideIdList.ChildElements[from] as SlideId;

            SlideId targetSlide;

            // Identify the position of the target slide after which to move the source slide
            if (to == 0)
            {
                targetSlide = null;
            }
            else if (from < to)
            {
                targetSlide = (SlideId)slideIdList.ChildElements[to];
            }
            else
            {
                targetSlide = (SlideId)slideIdList.ChildElements[to - 1];
            }

            // Remove the source slide from its current position.
            sourceSlide.Remove();

            // Insert the source slide at its new position after the target slide.
            slideIdList.InsertAfter(sourceSlide, targetSlide);

            // Save the modified presentation.
            presentation.Save();
        }

        public SCImage Background => this.backgroundImage.Value;

        public string CustomData
        {
            get => this.GetCustomData();
            set => this.SetCustomData(value);
        }

        public bool Hidden => this.SlidePart.Slide.Show != null && this.SlidePart.Slide.Show.Value == false;

        #endregion Public Properties

        public bool IsRemoved { get; set; }

        public SCPresentation ParentPresentation { get; }

        internal SlidePart SlidePart { get; }

        protected ResettableLazy<ShapeCollection> _shapes { get; }

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
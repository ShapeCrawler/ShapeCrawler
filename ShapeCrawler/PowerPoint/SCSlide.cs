using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using AngleSharp;
using AngleSharp.Css.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Factories;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Statics;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler
{
    internal class SCSlide : ISlide, IPresentationComponent
    {
        private readonly Lazy<SCImage> backgroundImage;
        private Lazy<CustomXmlPart> customXmlPart;
        private int? lastRid;

        internal SCSlide(SCPresentation parentPresentation, SlidePart slidePart, SlideId slideId)
        {
            this.PresentationInternal = parentPresentation;
            this.ParentPresentation = parentPresentation;
            this.SDKSlidePart = slidePart;
            this.shapes = new ResettableLazy<ShapeCollection>(() => ShapeCollection.ForSlide(this.SDKSlidePart, this));
            this.backgroundImage = new Lazy<SCImage>(() => SCImage.GetSlideBackgroundImageOrDefault(this));
            this.customXmlPart = new Lazy<CustomXmlPart>(this.GetSldCustomXmlPart);
            this.SlideId = slideId;
        }

        internal readonly SlideId SlideId;

        public ISlideLayout ParentSlideLayout =>
            ((SlideMasterCollection)this.PresentationInternal.SlideMasters).GetSlideLayoutBySlide(this);

        public IShapeCollection Shapes => this.shapes.Value;

        public int Number
        {
            get => this.GetNumber();
            set => this.SetNumber(value);
        }


        public SCImage Background => this.backgroundImage.Value;

        public string CustomData
        {
            get => this.GetCustomData();
            set => this.SetCustomData(value);
        }

        public bool Hidden => this.SDKSlidePart.Slide.Show != null && this.SDKSlidePart.Slide.Show.Value == false;

        public bool IsRemoved { get; set; }

        public IPresentation ParentPresentation { get; }

        public SlidePart SDKSlidePart { get; }

        private ResettableLazy<ShapeCollection> shapes { get; }

        /// <summary>
        ///     Saves slide scheme in PNG file.
        /// </summary>
        public void SaveScheme(string filePath)
        {
            SlideSchemeService.SaveScheme(this.shapes.Value, this.PresentationInternal.SlideWidth,
                this.PresentationInternal.SlideHeight, filePath);
        }

        public async Task<string> ToHtml()
        {
            var slideWidthPx = this.PresentationInternal.SlideWidth;
            var slideHeightPx = this.PresentationInternal.SlideHeight;

            var config = Configuration.Default.WithCss();
            var context = BrowsingContext.New(config);
            var document = await context.OpenNewAsync().ConfigureAwait(false);

            var styleBuilder = new StringBuilder();
            styleBuilder.Append("display: flex; ");
            styleBuilder.Append("justify-content: center; ");
            styleBuilder.Append("background: cadetblue; ");

            styleBuilder.Append($"width: {slideWidthPx}px; ");
            styleBuilder.Append($"height: {slideHeightPx}px; ");

            var main = document.CreateElement("main");
            main.SetStyle(styleBuilder.ToString());

            document.Body!.AppendChild(main);

            return document.DocumentElement.OuterHtml;
        }

        /// <summary>
        ///     Saves slide scheme in stream.
        /// </summary>
        public void SaveScheme(Stream stream)
        {
            SlideSchemeService.SaveScheme(this.shapes.Value, this.PresentationInternal.SlideWidth,
                this.PresentationInternal.SlideHeight, stream);
        }

        public void Hide()
        {
            if (this.SDKSlidePart.Slide.Show == null)
            {
                var showAttribute = new OpenXmlAttribute("show", string.Empty, "0");
                this.SDKSlidePart.Slide.SetAttribute(showAttribute);
            }
            else
            {
                this.SDKSlidePart.Slide.Show = false;
            }
        }

        #region Private Methods

        private int GetNumber()
        {
            var presentationPart = this.PresentationInternal.PresentationDocument.PresentationPart;
            string currentSlidePartId = presentationPart.GetIdOfPart(this.SDKSlidePart);
            List<SlideId> slideIdList =
                presentationPart.Presentation.SlideIdList.ChildElements.OfType<SlideId>().ToList();
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

            if (to < 0 || from >= this.PresentationInternal.Slides.Count || to == from)
            {
                throw new ArgumentOutOfRangeException(nameof(to));
            }

            PresentationPart presentationPart = this.PresentationInternal.PresentationDocument.PresentationPart;

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
                CustomXmlPart newSlideCustomXmlPart = this.SDKSlidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                customXmlPartStream = newSlideCustomXmlPart.GetStream();
#if NETSTANDARD2_0
                this.customXmlPart = new Lazy<CustomXmlPart>(() => newSlideCustomXmlPart);
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
            foreach (CustomXmlPart customXmlPart in this.SDKSlidePart.CustomXmlParts)
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

            this.PresentationInternal.ThrowIfClosed();
        }

        #endregion Private Methods

        public SCPresentation PresentationInternal { get; }

        internal string GenerateNextRelationshipId()
        {
            if (this.lastRid != null)
            {
                return $"rId{++this.lastRid}";
            }

            var idList = this.GetIdList();

            this.lastRid = idList.Max();
            return $"rId{++this.lastRid}";
        }

        private List<int> GetIdList()
        {
            var idList = new List<int>();

            foreach (var idPartPair in this.SDKSlidePart.Parts)
            {
                var matched = Regex.Match(idPartPair.RelationshipId, @"(?<=rId)\d+");
                var hasInt = int.TryParse(matched.Value, out var rIdInt);
                if (hasInt)
                {
                    idList.Add(rIdInt);
                }
            }

            foreach (var relationship in this.SDKSlidePart.HyperlinkRelationships)
            {
                var matched = Regex.Match(relationship.Id, @"(?<=rId)\d+");
                var hasInt = int.TryParse(matched.Value, out var rIdInt);
                if (hasInt)
                {
                    idList.Add(rIdInt);
                }
            }

            return idList;
        }
    }
}
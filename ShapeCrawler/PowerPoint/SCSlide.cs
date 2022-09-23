using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AngleSharp;
using AngleSharp.Css.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Services;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Statics;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents Slide.
    /// </summary>
    internal class SCSlide : SlideBase, ISlide, IPresentationComponent
    {
        internal readonly SlideId slideId;
        private readonly Lazy<SCImage> backgroundImage;
        private Lazy<CustomXmlPart> customXmlPart;
        private ResettableLazy<ShapeCollection> shapes;

        internal SCSlide(SCPresentation parentPresentation, SlidePart slidePart, SlideId slideId)
        {
            this.PresentationInternal = parentPresentation;
            this.Presentation = parentPresentation;
            this.SDKSlidePart = slidePart;
            this.shapes = new ResettableLazy<ShapeCollection>(() => ShapeCollection.ForSlide(this.SDKSlidePart, this));
            this.backgroundImage = new Lazy<SCImage>(() => SCImage.ForBackground(this));
            this.customXmlPart = new Lazy<CustomXmlPart>(this.GetSldCustomXmlPart);
            this.slideId = slideId;
        }

        public ISlideLayout SlideLayout => ((SlideMasterCollection)this.PresentationInternal.SlideMasters).GetSlideLayoutBySlide(this);

        public IShapeCollection Shapes => this.shapes.Value;

        public override bool IsRemoved { get; set; }
        

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

        public IPresentation Presentation { get; }

        public SlidePart SDKSlidePart { get; }

        public SCPresentation PresentationInternal { get; }
        
        internal SCSlideLayout SlideLayoutInternal => (SCSlideLayout)this.SlideLayout;

        internal override TypedOpenXmlPart TypedOpenXmlPart => this.SDKSlidePart;

        public override void ThrowIfRemoved()
        {
            if (this.IsRemoved)
            {
                throw new ElementIsRemovedException("Slide was removed");
            }

            this.PresentationInternal.ThrowIfClosed();
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

        public IList<ITextFrame> GetAllTextFrames()
        {
            List<ITextFrame> returnList = new List<ITextFrame>();

            // this will add all textboxes from shapes on that slide that directly inherit ITextBoxContainer
            returnList.AddRange(this.Shapes.OfType<ITextFrameContainer>()
                .Where(t => t.TextFrame != null)
                .Select(t => t.TextFrame)
                .ToList());

            // if this slide contains a table, the cells from that table will have to be added as well, since they inherit from ITextBoxContainer but are not direct descendants of the slide
            var tablesOnSlide = this.Shapes.OfType<ITable>().ToList();
            if (tablesOnSlide.Any())
            {
                returnList.AddRange(tablesOnSlide.SelectMany(table => table.Rows.SelectMany(row => row.Cells).Select(cell => cell.TextFrame)));
            }

            // if there are groups on that slide, they need to be added as well since those are not direct descendants of the slide either
            var groupsOnSlide = this.Shapes.OfType<IGroupShape>().ToList();
            if (groupsOnSlide.Any())
            {
                foreach (var group in groupsOnSlide)
                {
                    this.AddAllTextboxesInGroupToList(group, returnList);
                }
            }

            return returnList;
        }

        /// <summary>
        /// recursively iterate through a group and add all textboxes in that group to a list.
        /// </summary>
        /// <param name="group"></param>
        /// <param name="textBoxes"></param>
        private void AddAllTextboxesInGroupToList(IGroupShape group, List<ITextFrame> textBoxes)
        {
            foreach (var shape in group.Shapes)
            {
                switch (shape.ShapeType)
                {
                    case SCShapeType.GroupShape:
                        this.AddAllTextboxesInGroupToList((IGroupShape)shape, textBoxes);
                        break;
                    case SCShapeType.AutoShape:
                        if (shape is ITextFrameContainer)
                        {
                            textBoxes.Add(((ITextFrameContainer)shape).TextFrame);
                        }

                        break;
                    default:
                        break;
                }
            }
        }

        private int GetNumber()
        {
            var presentationPart = this.PresentationInternal.SdkPresentation.PresentationPart;
            string currentSlidePartId = presentationPart.GetIdOfPart(this.SDKSlidePart);
            List<SlideId> slideIdList = presentationPart.Presentation.SlideIdList.ChildElements.OfType<SlideId>().ToList();
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

            PresentationPart presentationPart = this.PresentationInternal.SdkPresentation.PresentationPart;

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
    }
}
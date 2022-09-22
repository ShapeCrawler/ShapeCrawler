using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Services;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMasters
{
    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC — ShapeCrawler")]
    internal class SCSlideMaster : SlideBase, ISlideMaster
    {
        private readonly ResettableLazy<List<SCSlideLayout>> slideLayouts;

        internal SCSlideMaster(SCPresentation pres, P.SlideMaster pSlideMaster)
        {
            this.Presentation = pres;
            this.PSlideMaster = pSlideMaster;
            this.slideLayouts = new ResettableLazy<List<SCSlideLayout>>(this.GetSlideLayouts);
        }
        
        public SCImage Background => this.GetBackground();

        public IReadOnlyList<ISlideLayout> SlideLayouts => this.slideLayouts.Value;

        public IShapeCollection Shapes => ShapeCollection.ForSlideLayout(this.PSlideMaster.CommonSlideData.ShapeTree, this);

        public override bool IsRemoved { get; set; }
        
        internal P.SlideMaster PSlideMaster { get; }

        internal SCPresentation Presentation { get; }

        internal Dictionary<int, FontData> BodyParaLvlToFontData =>
            FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles!.BodyStyle!);

        internal Dictionary<int, FontData> TitleParaLvlToFontData =>
            FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles!.TitleStyle!);

        internal ThemePart ThemePart => this.PSlideMaster.SlideMasterPart.ThemePart;

        internal ShapeCollection ShapesInternal => (ShapeCollection)this.Shapes;

        internal override TypedOpenXmlPart TypedOpenXmlPart => this.PSlideMaster.SlideMasterPart!;

        public override void ThrowIfRemoved()
        {
            if (this.IsRemoved)
            {
                throw new ElementIsRemovedException("Slide MAster is removed");
            }
            
            this.Presentation.ThrowIfClosed();
        }
        
        internal bool TryGetFontSizeFromBody(int paragraphLvl, out int fontSize)
        {
            Dictionary<int, FontData> bodyParaLvlToFontData =
                FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles.BodyStyle);
            if (bodyParaLvlToFontData.TryGetValue(paragraphLvl, out FontData fontData))
            {
                if (fontData.FontSize != null)
                {
                    fontSize = fontData.FontSize;
                    return true;
                }
            }

            fontSize = -1;
            return false;
        }

        internal bool TryGetFontSizeFromOther(int paragraphLvl, out int fontSize)
        {
            DocumentFormat.OpenXml.Presentation.TextStyles pTextStyles = this.PSlideMaster.TextStyles;

            // Other
            Dictionary<int, FontData> otherStyleLvlToFontData =
                FontDataParser.FromCompositeElement(pTextStyles.OtherStyle);
            if (otherStyleLvlToFontData.ContainsKey(paragraphLvl))
            {
                if (otherStyleLvlToFontData[paragraphLvl].FontSize != null)
                {
                    fontSize = otherStyleLvlToFontData[paragraphLvl].FontSize;
                    return true;
                }
            }

            fontSize = -1;
            return false;
        }
        
        private SCImage GetBackground()
        {
            return null;
        }
        
        private List<SCSlideLayout> GetSlideLayouts()
        {
            IEnumerable<SlideLayoutPart> sldLayoutParts = this.PSlideMaster.SlideMasterPart.SlideLayoutParts;
            var slideLayouts = new List<SCSlideLayout>(sldLayoutParts.Count());
            foreach (SlideLayoutPart sldLayoutPart in sldLayoutParts)
            {
                slideLayouts.Add(new SCSlideLayout(this, sldLayoutPart));
            }

            return slideLayouts;
        }
    }
}
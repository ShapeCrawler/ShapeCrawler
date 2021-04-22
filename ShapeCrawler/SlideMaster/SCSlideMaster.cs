using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Drawing;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMaster;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace

namespace ShapeCrawler
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class SCSlideMaster : IBaseSlide //TODO: add ISlideMaster interface and make internal
    {
        private readonly ResettableLazy<List<SCSlideLayout>> _sldLayouts;
        internal readonly P.SlideMaster PSlideMaster;

        internal SCSlideMaster(SCPresentation presentation, P.SlideMaster pSlideMaster)
        {
            Presentation = presentation;
            PSlideMaster = pSlideMaster;
            _sldLayouts = new ResettableLazy<List<SCSlideLayout>>(GetSlideLayouts);
        }

        internal SCPresentation Presentation { get; }

        internal Dictionary<int, FontData> BodyParaLvlToFontData =>
            FontDataParser.FromCompositeElement(PSlideMaster.TextStyles.BodyStyle);

        internal Dictionary<int, FontData> TitleParaLvlToFontData =>
            FontDataParser.FromCompositeElement(PSlideMaster.TextStyles.TitleStyle);

        public void Hide() //TODO: does it need?
        {
            throw new NotImplementedException();
        }

        private List<SCSlideLayout> GetSlideLayouts()
        {
            IEnumerable<SlideLayoutPart> sldLayoutParts = PSlideMaster.SlideMasterPart.SlideLayoutParts;
            var slideLayouts = new List<SCSlideLayout>(sldLayoutParts.Count());
            foreach (SlideLayoutPart sldLayoutPart in sldLayoutParts)
            {
                slideLayouts.Add(new SCSlideLayout(this, sldLayoutPart));
            }

            return slideLayouts;
        }

        internal bool TryGetFontSizeFromBody(int paragraphLvl, out int fontSize)
        {
            Dictionary<int, FontData> bodyParaLvlToFontData =
                FontDataParser.FromCompositeElement(PSlideMaster.TextStyles.BodyStyle);
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

        internal A.SchemeColorValues GetFontColorHexFromBody(int paragraphLvl)
        {
            Dictionary<int, FontData> bodyParaLvlToFontData =
                FontDataParser.FromCompositeElement(PSlideMaster.TextStyles.BodyStyle);

            return bodyParaLvlToFontData[paragraphLvl].ASchemeColor.Val;
        }

        internal bool TryGetFontSizeFromOther(int paragraphLvl, out int fontSize)
        {
            P.TextStyles pTextStyles = PSlideMaster.TextStyles;

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

        #region Public Properties

        public ShapeCollection Shapes => ShapeCollection.CreateForSlideMaster(this);
        public int Number { get; } //TODO: does it need?
        public SCImage Background { get; }
        public string CustomData { get; set; } //TODO: does it need?
        public bool Hidden { get; } //TODO: does it need?
        public IReadOnlyList<SCSlideLayout> SlideLayouts => _sldLayouts.Value;

        #endregion Public Properties
    }
}
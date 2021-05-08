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
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a Slide Master.
    /// </summary>
    public interface ISlideMaster : IBaseSlide
    {
        /// <summary>
        ///     Gets background image.
        /// </summary>
        SCImage Background { get; }

        /// <summary>
        ///     Gets collection of Slide Layouts.
        /// </summary>
        IReadOnlyList<ISlideLayout> SlideLayouts { get; }
    }

    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC — ShapeCrawler")]
    internal class SCSlideMaster : ISlideMaster // TODO: make internal
    {
        private readonly ResettableLazy<List<SCSlideLayout>> slideLayouts;
        internal readonly P.SlideMaster PSlideMaster;

        internal SCSlideMaster(SCPresentation parentPresentation, P.SlideMaster pSlideMaster)
        {
            this.ParentPresentation = parentPresentation;
            this.PSlideMaster = pSlideMaster;
            this.slideLayouts = new ResettableLazy<List<SCSlideLayout>>(this.GetSlideLayouts);
        }

        internal SCPresentation ParentPresentation { get; }

        internal Dictionary<int, FontData> BodyParaLvlToFontData =>
            FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles.BodyStyle);

        internal Dictionary<int, FontData> TitleParaLvlToFontData =>
            FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles.TitleStyle);

        internal ThemePart ThemePart => this.PSlideMaster.SlideMasterPart.ThemePart;

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
            P.TextStyles pTextStyles = this.PSlideMaster.TextStyles;

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

        public void ThrowIfRemoved()
        {
            throw new NotImplementedException();
        }

        #region Public Properties

        public SCImage Background => throw new NotImplementedException();

        public IReadOnlyList<ISlideLayout> SlideLayouts => this.slideLayouts.Value;

        IShapeCollection IBaseSlide.Shapes => ShapeCollection.CreateForSlideMaster(this);

        #endregion Public Properties
    }
}
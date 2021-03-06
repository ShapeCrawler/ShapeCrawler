using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMaster
{
    public class SlideMasterSc : ISlide //TODO: add ISlideMaster interface
    {
        private readonly ResettableLazy<List<SlideLayoutSc>> _sldLayouts;
        internal readonly P.SlideMaster PSlideMaster;

        internal SlideMasterSc(PresentationSc presentation, P.SlideMaster pSlideMaster)
        {
            Presentation = presentation;
            PSlideMaster = pSlideMaster;
            _sldLayouts = new ResettableLazy<List<SlideLayoutSc>>(() => GetSlideLayouts());
        }

        internal PresentationSc Presentation { get; }

        public void Hide() //TODO: does it need?
        {
            throw new NotImplementedException();
        }

        private List<SlideLayoutSc> GetSlideLayouts()
        {
            IEnumerable<SlideLayoutPart> sldLayoutParts = PSlideMaster.SlideMasterPart.SlideLayoutParts;
            var slideLayouts = new List<SlideLayoutSc>(sldLayoutParts.Count());
            foreach (SlideLayoutPart sldLayoutPart in sldLayoutParts)
            {
                slideLayouts.Add(new SlideLayoutSc(this, sldLayoutPart));
            }

            return slideLayouts;
        }

        #region Public Properties

        public ShapeCollection Shapes => ShapeCollection.CreateForSlideMaster(this);
        public int Number { get; } //TODO: does it need?
        public ImageSc Background { get; }
        public string CustomData { get; set; } //TODO: does it need?
        public bool Hidden { get; } //TODO: does it need?
        public IReadOnlyList<SlideLayoutSc> SlideLayouts => _sldLayouts.Value;

        #endregion Public Properties
    }
}
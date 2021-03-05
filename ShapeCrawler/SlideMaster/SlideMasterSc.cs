using System;
using System.Collections;
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
        private readonly P.SlideMaster _pSlideMaster;
        private readonly ResettableLazy<List<SlideLayoutSc>> _sldLayouts;

        internal PresentationSc Presentation { get; }

        internal SlideMasterSc(PresentationSc presentation, P.SlideMaster pSlideMaster)
        {
            Presentation = presentation;
            _pSlideMaster = pSlideMaster;
            _sldLayouts = new ResettableLazy<List<SlideLayoutSc>>(() => GetSlideLayouts());
        }

        #region Public Properties

        public ShapeCollection Shapes { get; }
        public int Number { get; } //TODO: does it need?
        public ImageSc Background { get; }
        public string CustomData { get; set; } //TODO: does it need?
        public bool Hidden { get; } //TODO: does it need?
        public IReadOnlyList<SlideLayoutSc> SlideLayouts => _sldLayouts.Value;

        private List<SlideLayoutSc> GetSlideLayouts()
        {
            IEnumerable<SlideLayoutPart> sldLayoutParts = _pSlideMaster.SlideMasterPart.SlideLayoutParts;
            var slideLayouts = new List<SlideLayoutSc>(sldLayoutParts.Count());
            foreach (SlideLayoutPart sldLayoutPart in sldLayoutParts)
            {
                slideLayouts.Add(new SlideLayoutSc(this, sldLayoutPart));
            }

            return slideLayouts;
        }

        #endregion Public Properties

        public void Hide() //TODO: does it need?
        {
            throw new NotImplementedException();
        }
    }
}
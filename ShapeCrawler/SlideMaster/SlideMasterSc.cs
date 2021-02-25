using System;
using ShapeCrawler.Collections;
using ShapeCrawler.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMaster
{
    public class SlideMasterSc : ISlide
    {
        private readonly P.SlideMaster _pSlideMaster;

        public SlideMasterSc(P.SlideMaster pSlideMaster)
        {
            _pSlideMaster = pSlideMaster;
        }

        ShapeCollection ISlide.Shapes { get; }

        #region Public Properties

        public MasterShapeCollection Shapes =>
            ShapeCollection.CreateForMasterSlide(this, _pSlideMaster.CommonSlideData.ShapeTree);

        public int Number { get; }
        public ImageSc Background { get; }
        public string CustomData { get; set; }
        public bool Hidden { get; }

        public void Hide()
        {
            throw new NotImplementedException();
        }

        #endregion Public Properties
    }
}
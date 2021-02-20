using ShapeCrawler.Collections;
using ShapeCrawler.Factories.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMaster
{
    public class SlideMasterSc : ISlide
    {
        private readonly P.SlideMaster _pSlideMaster;
        private ShapesCollection _shapes;

        public SlideMasterSc(P.SlideMaster pSlideMaster)
        {
            _pSlideMaster = pSlideMaster;
        }

        #region Public Properties

        public MasterShapesCollection Shapes =>
            ShapesCollection.CreateForMasterSlide(this, _pSlideMaster.CommonSlideData.ShapeTree);

        public int Number { get; }
        public ImageSc Background { get; }
        public string CustomData { get; set; }
        public bool Hidden { get; }
        public void Hide()
        {
            throw new System.NotImplementedException();
        }

        #endregion Public Properties

        ShapesCollection ISlide.Shapes => _shapes;
    }
}
using ShapeCrawler.Collections;
using ShapeCrawler.Models;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMaster
{
    public class SlideMasterSc : ISlide
    {
        private readonly P.SlideMaster _pSlideMaster;

        #region Public Properties

        public MasterShapesCollection Shapes => ShapesCollection.CreateForMasterSlide(this, _pSlideMaster.CommonSlideData.ShapeTree);

        #endregion Public Properties

        public SlideMasterSc(P.SlideMaster pSlideMaster)
        {
            _pSlideMaster = pSlideMaster;
        }
    }
}
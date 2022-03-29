using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing
{
    internal class MasterPicture : MasterShape, IPicture
    {
        private readonly StringValue picReference;

        internal MasterPicture(P.Picture pPicture, SCSlideMaster slideMaster, StringValue picReference)
            : base(pPicture, slideMaster)
        {
            this.picReference = picReference;
            this.PresentationInternal = slideMaster.ParentPresentation;
        }

        public SCImage Image => this.GetImage();

        public ShapeType ShapeType => ShapeType.Picture;

        public override SCPresentation PresentationInternal { get; }

        private SCImage GetImage()
        {
            var sldMasterPart = this.SlideMasterInternal.PSlideMaster.SlideMasterPart;
            var imagePart = (ImagePart)sldMasterPart.GetPartById(picReference.Value);

            return new SCImage(imagePart, this, picReference, sldMasterPart);
        }
    }
}
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing
{
    internal class LayoutPicture : LayoutShape, IPicture
    {
        private readonly StringValue picReference;

        internal LayoutPicture(P.Picture pPicture, SCSlideLayout layout, StringValue picReference)
            : base(layout, pPicture)
        {
            this.picReference = picReference;
        }

        public SCImage Image => this.GetImage();

        public SCShapeType ShapeType => SCShapeType.Picture;

        private SCImage GetImage()
        {
            var imagePart = (ImagePart)this.SlideLayoutInternal.SlideLayoutPart.GetPartById(this.picReference.Value);

            return SCImage.Create(imagePart, this, this.picReference, this.SlideLayoutInternal.SlideLayoutPart);
        }
    }
}
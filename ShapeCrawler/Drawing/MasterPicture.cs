using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal class MasterPicture : MasterShape, IPicture
{
    private readonly StringValue picReference;

    internal MasterPicture(P.Picture pPicture, SCSlideMaster slideMaster, StringValue picReference)
        : base(pPicture, slideMaster)
    {
        this.picReference = picReference;
    }

    public IImage Image => this.GetImage();

    public override SCShapeType ShapeType => SCShapeType.Picture;

    private SCImage GetImage()
    {
        var sldMasterPart = this.SlideMasterInternal.PSlideMaster.SlideMasterPart!;
        var imagePart = (ImagePart)sldMasterPart.GetPartById(this.picReference.Value!);

        return SCImage.Create(imagePart, this, this.picReference, sldMasterPart);
    }
}
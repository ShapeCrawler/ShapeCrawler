using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Presentations;

internal readonly ref struct SCSlideMasterPart
{
    private readonly SlideMasterPart slideMasterPart;

    internal SCSlideMasterPart(SlideMasterPart slideMasterPart)
    {
        this.slideMasterPart = slideMasterPart;
    }

    internal void RemoveLayoutsExcept(SlideLayoutPart exceptSlideLayoutPart)
    {
        var pSlideLayoutIds = this.slideMasterPart.SlideMaster.SlideLayoutIdList!.OfType<DocumentFormat.OpenXml.Presentation.SlideLayoutId>();
        foreach (var slideLayoutPart in this.slideMasterPart.SlideLayoutParts.ToList())
        {
            if (slideLayoutPart == exceptSlideLayoutPart)
            {
                continue;
            }

            var id = this.slideMasterPart.GetIdOfPart(slideLayoutPart);
            var layoutId = pSlideLayoutIds.First(x => x.RelationshipId == id);
            layoutId.Remove();
            this.slideMasterPart.DeletePart(slideLayoutPart);
        }
    }
}
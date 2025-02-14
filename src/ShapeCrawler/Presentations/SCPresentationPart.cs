using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Presentations;

internal readonly ref struct SCPresentationPart
{
    private readonly PresentationPart presentationPart;

    internal SCPresentationPart(PresentationPart presentationPart)
    {
        this.presentationPart = presentationPart;
    }

    internal void AddSlidePart(SlidePart slidePart)
    {
        var rId = new SCOpenXmlPart(this.presentationPart).NextRelationshipId();
        var addedSlidePart = this.presentationPart.AddPart(slidePart, rId);

        var notesSlidePartAddedSlidePart = addedSlidePart.GetPartsOfType<NotesSlidePart>().FirstOrDefault();
        notesSlidePartAddedSlidePart?.DeletePart(notesSlidePartAddedSlidePart.NotesMasterPart!);

        rId = new SCOpenXmlPart(this.presentationPart).NextRelationshipId();
        var addedSlideMasterPart = this.presentationPart.AddPart(addedSlidePart.SlideLayoutPart!.SlideMasterPart!, rId);
        var layoutIdList = addedSlideMasterPart.SlideMaster.SlideLayoutIdList!.OfType<P.SlideLayoutId>();
        foreach (var layoutId in layoutIdList.ToList())
        {
            if (!addedSlideMasterPart.TryGetPartById(layoutId.RelationshipId!, out _))
            {
                layoutId.Remove();
            }
        }
    }

    internal T Last<T>()
        where T : OpenXmlPart
        => this.presentationPart.GetPartsOfType<T>().Last();
}
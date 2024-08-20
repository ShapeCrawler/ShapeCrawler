using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;

namespace ShapeCrawler;

internal readonly ref struct WrappedPresentationPart
{
    private readonly PresentationPart presentationPart;

    internal WrappedPresentationPart(PresentationPart presentationPart)
    {
        this.presentationPart = presentationPart;
    }
    
    internal void AddSlidePart(SlidePart sourceSlidePart)
    {
        var rId = this.presentationPart.NextRelationshipId();
        var addedSlidePart = this.presentationPart.AddPart(sourceSlidePart, rId);

        var notesSlidePartAddedSlidePart = addedSlidePart.GetPartsOfType<NotesSlidePart>().FirstOrDefault();
        if (notesSlidePartAddedSlidePart != null)
        {
            addedSlidePart.DeletePart(notesSlidePartAddedSlidePart);
        }

        rId = this.presentationPart.NextRelationshipId();
        var addedSlideMasterPart = this.presentationPart.AddPart(addedSlidePart.SlideLayoutPart!.SlideMasterPart!, rId);
        var layoutIdList = addedSlideMasterPart.SlideMaster.SlideLayoutIdList!.OfType<DocumentFormat.OpenXml.Presentation.SlideLayoutId>();
        foreach (var layoutId in layoutIdList.ToList())
        {
            if (!addedSlideMasterPart.TryGetPartById(layoutId.RelationshipId!, out _))
            {
                layoutId.Remove();
            }
        }
    }

    internal T Last<T>() where T : OpenXmlPart => this.presentationPart.GetPartsOfType<T>().Last();
}
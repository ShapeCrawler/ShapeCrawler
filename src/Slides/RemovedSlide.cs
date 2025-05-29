using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Presentations;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

#if DEBUG
#endif

namespace ShapeCrawler;

internal sealed class RemovedSlide : Slide
{
    internal RemovedSlide(ISlideLayout slideLayout, ISlideShapeCollection shapes, SlidePart slidePart)
        : base(slideLayout, shapes, slidePart)
    {
    }

    public override void Remove()
    {
        var presDocument = (PresentationDocument)this.SlidePart.OpenXmlPackage;
        var presPart = presDocument.PresentationPart!;
        var pPresentation = presDocument.PresentationPart!.Presentation;
        var slideIdList = pPresentation.SlideIdList!;

        // Find the exact SlideId corresponding to this slide
        var slideIdRelationship = presPart.GetIdOfPart(this.SlidePart);
        var removingPSlideId = slideIdList.Elements<P.SlideId>()
                                   .FirstOrDefault(slideId => slideId.RelationshipId!.Value == slideIdRelationship) ??
                               throw new SCException("Could not find slide ID in presentation.");

        // Handle section references
        var sectionList = pPresentation.PresentationExtensionList?.Descendants<P14.SectionList>().FirstOrDefault();
        var removingSectionSlideIdListEntry = sectionList?.Descendants<P14.SectionSlideIdListEntry>()
            .FirstOrDefault(s => s.Id! == removingPSlideId.Id!);
        removingSectionSlideIdListEntry?.Remove();
        
        slideIdList.RemoveChild(removingPSlideId);
        pPresentation.Save();

        var removingSlideIdRelationshipId = removingPSlideId.RelationshipId!;
        new SCPPresentation(pPresentation).RemoveSlideIdFromCustomShow(removingSlideIdRelationshipId.Value!);

        var removingSlidePart = (SlidePart)presPart.GetPartById(removingSlideIdRelationshipId!);
        presPart.DeletePart(removingSlidePart);

        // Final save to ensure structure is consistent
        presPart.Presentation.Save();
    }
}
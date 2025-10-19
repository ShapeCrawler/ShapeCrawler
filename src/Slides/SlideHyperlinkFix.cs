using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

/// <summary>
/// Adjusts slide hyperlinks so they target the correct relationships in a destination presentation.
/// </summary>
internal static class SlideHyperlinkFix
{
    internal static void FixSlideHyperlinks(
        SlidePart sourceSlidePart,
        SlidePart clonedSlidePart,
        PresentationPart targetPresentationPart)
    {
        var sourcePresentation = ((PresentationDocument)sourceSlidePart.OpenXmlPackage).PresentationPart!;

        var hyperlinks = clonedSlidePart.Slide.Descendants<HyperlinkOnClick>()
            .Where(h => h.Action?.Value == "ppaction://hlinksldjump" && !string.IsNullOrEmpty(h.Id?.Value));

        foreach (var hyperlink in hyperlinks)
        {
            try
            {
                var sourceTargetSlidePart = (SlidePart)sourceSlidePart.GetPartById(hyperlink.Id!.Value!);
                var sourceSlideIdList = sourcePresentation.Presentation.SlideIdList!.ChildElements.OfType<SlideId>();
                var sourceTargetSlideRelId = sourcePresentation.GetIdOfPart(sourceTargetSlidePart);

                var sourceSlideNumber = 0;
                foreach (var slideId in sourceSlideIdList)
                {
                    sourceSlideNumber++;
                    if (slideId.RelationshipId == sourceTargetSlideRelId)
                    {
                        break;
                    }
                }

                var targetSlideIdList =
                    targetPresentationPart.Presentation.SlideIdList!.ChildElements.OfType<SlideId>();
                var targetSlideId = targetSlideIdList.ElementAtOrDefault(sourceSlideNumber - 1);

                if (targetSlideId != null)
                {
                    var targetSlidePart = (SlidePart)targetPresentationPart.GetPartById(targetSlideId.RelationshipId!);
                    var newRelationship = clonedSlidePart.AddPart(targetSlidePart);
                    var newRelId = clonedSlidePart.GetIdOfPart(newRelationship);
                    hyperlink.Id = newRelId;
                }
            }
            catch
            {
                hyperlink.Id = null;
                hyperlink.Action = null;
            }
        }
    }
}

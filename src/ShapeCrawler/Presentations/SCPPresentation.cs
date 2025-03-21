using System.Linq;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Presentations;

// ReSharper disable once InconsistentNaming
internal readonly ref struct SCPPresentation(P.Presentation pPresentation)
{
    internal void RemoveSlideIdFromCustomShow(string slideIdRelationshipId)
    {
        if (pPresentation.CustomShowList == null)
        {
            return;
        }

        foreach (var pCustomShow in pPresentation.CustomShowList.Elements<P.CustomShow>())
        {
            pCustomShow.SlideList?
                .Elements<P.SlideListEntry>()
                .Where(entry => entry.Id == slideIdRelationshipId)
                .ToList()
                .ForEach(entry => pCustomShow.SlideList.RemoveChild(entry));
        }
    }
}
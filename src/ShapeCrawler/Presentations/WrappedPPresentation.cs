using System.Linq;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Presentations;

internal readonly ref struct WrappedPPresentation
{
    private readonly P.Presentation pPresentation;

    internal WrappedPPresentation(P.Presentation pPresentation)
    {
        this.pPresentation = pPresentation;
    }

    internal void RemoveSlideIdFromCustomShow(string slideIdRelationshipId)
    {
        if (this.pPresentation.CustomShowList == null)
        {
            return;
        }

        foreach (var pCustomShow in this.pPresentation.CustomShowList.Elements<P.CustomShow>())
        {
            pCustomShow.SlideList?
                .Elements<P.SlideListEntry>()
                .Where(entry => entry.Id == slideIdRelationshipId)
                .ToList()
                .ForEach(entry => pCustomShow.SlideList.RemoveChild(entry));
        }
    }
}
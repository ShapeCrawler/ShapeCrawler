using System.Linq;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Presentations;

// ReSharper disable once InconsistentNaming
internal readonly ref struct SCPPresentation
{
    private readonly P.Presentation pPresentation;

    internal SCPPresentation(P.Presentation pPresentation)
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
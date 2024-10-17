using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

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

        foreach (var customShow in pPresentation.CustomShowList.Elements<P.CustomShow>())
        {
            customShow.SlideList?
                .Elements<P.SlideListEntry>()
                .Where(entry => entry.Id == slideIdRelationshipId)
                .ToList()
                .ForEach(entry => customShow.SlideList.RemoveChild(entry));
        }
    }
}
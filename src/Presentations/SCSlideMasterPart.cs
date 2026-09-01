using System.IO;
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

    /// <summary>
    /// Adds a copy of the specified layout and its relationships to the wrapped slide master part.
    /// </summary>
    /// <param name="sourceLayoutPart">Source layout part.</param>
    /// <returns>Added layout part.</returns>
    internal SlideLayoutPart AddLayout(SlideLayoutPart sourceLayoutPart)
    {
        var targetLayoutPart = this.slideMasterPart.AddNewPart<SlideLayoutPart>();
        using var sourceStream = sourceLayoutPart.GetStream();
        sourceStream.Position = 0;
        using var destinationStream = targetLayoutPart.GetStream(FileMode.Create, FileAccess.Write);
        sourceStream.CopyTo(destinationStream);

        foreach (var childPart in sourceLayoutPart.Parts)
        {
            var targetChildPart = childPart.OpenXmlPart is SlideMasterPart
                ? this.slideMasterPart
                : childPart.OpenXmlPart;
            targetLayoutPart.AddPart(targetChildPart, childPart.RelationshipId);
        }

        return targetLayoutPart;
    }

    internal void RemoveLayoutsExcept(SlideLayoutPart exceptSlideLayoutPart)
    {
        var pSlideLayoutIds = this.slideMasterPart.SlideMaster!.SlideLayoutIdList!.OfType<DocumentFormat.OpenXml.Presentation.SlideLayoutId>();
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
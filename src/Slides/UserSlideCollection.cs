using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class UserSlideCollection(IEnumerable<SlidePart> slideParts) : IReadOnlyList<IUserSlide>
{
    public int Count => this.GetSlides().Count();

    public IUserSlide this[int index] => this.GetSlides().ElementAt(index);

    public IEnumerator<IUserSlide> GetEnumerator() => this.GetSlides().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private IEnumerable<UserSlide> GetSlides()
    {
        if (!slideParts.Any())
        {
            yield break;
        }

        var presDocument = (PresentationDocument)slideParts.First().OpenXmlPackage;
        var presPart = presDocument.PresentationPart!;
        var pSlideIdList = presPart.Presentation.SlideIdList!.ChildElements.OfType<P.SlideId>();
        foreach (var pSlideId in pSlideIdList)
        {
            var slidePart = (SlidePart)presPart.GetPartById(pSlideId.RelationshipId!);
            var presImageFiles = new PresentationImageFiles(slideParts);
            yield return new DrawingSlide(
                new LayoutSlide(slidePart.SlideLayoutPart!),
                new UserSlideShapeCollection(
                    new ShapeCollection(slidePart),
                    new PictureShapeCollection(slidePart, presImageFiles),
                    new AudioVideoShapeCollection(slidePart, presImageFiles),
                    new ChartShapeCollection(slidePart),
                    slidePart),
                slidePart
            );
        }
    }
}
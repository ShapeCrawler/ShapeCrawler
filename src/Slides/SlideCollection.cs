using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class SlideCollection(IEnumerable<SlidePart> slideParts) : IReadOnlyList<ISlide>
{
    public int Count => this.GetSlides().Count();

    public ISlide this[int index] => this.GetSlides().ElementAt(index);

    public IEnumerator<ISlide> GetEnumerator() => this.GetSlides().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private IEnumerable<Slide> GetSlides()
    {
        if (!slideParts.Any())
        {
            yield break;
        }

        var presDocument = (PresentationDocument)slideParts.First().OpenXmlPackage;
        var presPart = presDocument.PresentationPart!;
        var pSlideIdList = presPart.Presentation.SlideIdList!.ChildElements.OfType<P.SlideId>().ToArray();
        foreach (var pSlideId in pSlideIdList)
        {
            var slidePart = (SlidePart)presPart.GetPartById(pSlideId.RelationshipId!);
            var presImageFiles = new PresentationImageFiles(slideParts);
            yield return new Slide(
                new SlideLayout(slidePart.SlideLayoutPart!),
                new SlideShapeCollection(
                    new ChartCollection(
                        new AudioVideoShapeCollection(
                            new PictureCollection(new ShapeCollection(slidePart), presImageFiles, slidePart),
                            presImageFiles, 
                            slidePart
                        ),
                        slidePart
                    ),
                    slidePart
                ),
                slidePart
            );
        }
    }
}